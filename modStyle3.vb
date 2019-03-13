Option Compare Text

Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO

Module modStyle3

    Sub AssignedBackCalcCalibr_3(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

        Dim boolOC As Boolean = False 'bool if eliminated
        Dim numNomConc As Decimal
        Dim var1, var2, var3, var4, var5, var10
        Dim dvDo As System.Data.DataView
        Dim strTName As String
        Dim strTNameO As String 'original
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
        Dim boolPro As Boolean

        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(20) 'FlagPercent array
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
        Dim arr1(1)
        Dim boolJustTable As Boolean

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short

        Dim intGroup As Short
        Dim intAnalyteID As Int64
        Dim strAnal As String
        Dim strAnalC As String
        Dim strMatrix As String

        Dim v1, v2, vU

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        With wd

            intTableID = 3

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
            strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies
            rowsX = tbl2.Select(strF)
            'If rowsX.Length = 0 Then
            '    strM = "Creating Summary of " & strTempInfo & " Back-Calculated Calibration Standard Concentrations Table ...."
            '    frmH.lblProgress.Text = strM
            '    frmH.Refresh()
            '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
            '    GoTo end2
            'End If

            'pull out nominal concentrations
            'first get assayid's
            If boolUseGroups Then
                strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR & " AND ASSAYID IS NOT NULL AND INTGROUP > 0"
            Else
                strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idTR & " AND ASSAYID IS NOT NULL"
            End If
            strS = "ASSAYID ASC"
            Dim dvNC1 As System.Data.DataView = New DataView(tbl2)
            Dim strFRuns As String
            strFRuns = strF
            dvNC1.RowFilter = strF
            dvNC1.Sort = strS
            Dim tblNC1 As System.Data.DataTable = dvNC1.ToTable("a", True, "ASSAYID")
            int1 = tblNC1.Rows.Count
            If int1 = 0 Then
                strM = "Creating " & strTName & "...."
                strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                frmH.lblProgress.Text = strM
                frmH.Refresh()
                'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")

                'don't say justtable yet
                'go further to evaluate
                'boolJustTable = True
                strF = "ASSAYID = 0"
                GoTo end3

            End If
            'assign assayids to string
            str1 = "ASSAYID = "
            For Count1 = 0 To int1 - 1
                var1 = tblNC1.Rows(Count1).Item("ASSAYID")
                If Count1 = int1 - 1 Then
                    str1 = str1 & var1
                Else
                    str1 = str1 & var1 & " OR ASSAYID = "
                End If
            Next



            'get runs for possible error message
            strFRuns = strFRuns & " AND (" & str1 & ")"
            dvNC1.RowFilter = strFRuns
            Dim tblNCRuns As System.Data.DataTable = dvNC1.ToTable("runs", True, "RUNID")
            Dim intRuns As Short
            Dim strRuns As String
            intRuns = tblNC1.Rows.Count
            str2 = "Run ID(s): "
            For Count1 = 0 To int1 - 1
                var1 = tblNCRuns.Rows(Count1).Item("RUNID")
                If Count1 = int1 - 1 Then
                    str2 = str2 & var1
                Else
                    str2 = str2 & var1 & ", "
                End If
            Next
            strRuns = str2

            strF = str1


            'MASTERASSAYID()
            'ANALYTEINDEX()
            'LEVELNUMBER()
            'CONCENTRATION()
            'STUDYID()
            'ASSAYID()

end3:


            'now filter tblBCStdsConc for assayids
            Dim dvNC3 As System.Data.DataView = New DataView(tblBCStdConcs)
            'dim dvNC3 as system.data.dataview = New DataView(tbl2)
            dvNC3.RowFilter = strF
            int1 = dvNC3.Count

            strF = "IsIntStd = 'No'"
            'for some reason, if sort isn't applied, then rows11 sort gets goofed up
            'it should sort in the order of the underlying table
            'strS = ReturnSort(False)
            strS = "INTORDER ASC, IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            Dim strMsg As String
            strMsg = ""
            For Count1 = 1 To intRowsAnal

                boolJustTable = False

                Dim arrLegend(4, 20)

                strTName = strTNameO

                intGroup = rows11(Count1 - 1).Item("INTGROUP")

                ctLegend = 0


                'If int1 = 0 Then
                '    boolJustTable = True
                '    strMsg = "There was a problem preparing this table."
                '    strMsg = strMsg & ChrW(10) & ChrW(10)
                '    strMsg = strMsg & "It is possible the assigned data for this table comes from an unaccepted analytical run(s):"
                '    strMsg = strMsg & ChrW(10) & ChrW(10)
                '    strMsg = strMsg & strRuns
                '    strMsg = strMsg & ChrW(10) & ChrW(10)
                '    strMsg = strMsg & "Please inspect the sample assignment in the Assigned Samples window."
                '    GoTo end1
                'End If

                Dim tblNC3 As System.Data.DataTable ' = dvNC3.ToTable()


                'for legend stuff
                intExp = 0

                'var1 = tblNC3.Rows.Count 'debugging

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
                strMatrix = rows11(Count1 - 1).Item("MATRIX")
                strAnal = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                strAnalC = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")

                If UseAnalyte(CStr(strDo)) Then
                Else
                    'NO!! No page return here
                    ''page setup according to configuration
                    'str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    ''insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    'Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                    GoTo next1
                End If

                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column

                Dim strM1 As String
                If bool Then 'continue

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)


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

                    'a table entry will be recorded in SRSummaryOfBCSC
                    ''enter a table record in tblTableN
                    'ctTableN = ctTableN + 1
                    'Dim dtblr As DataRow = tblTableN.NewRow
                    'dtblr.BeginEdit()
                    'dtblr.Item("TableNumber") = ctTableN
                    'dtblr.Item("AnalyteName") = strDo 'arrAnalytes(1, Count1)
                    'dtblr.Item("TableName") = strTNameO
                    'tblTableN.Rows.Add(dtblr)

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

                    'Legend:
                    'tbl1 = tblAnalysisResultsHome
                    'tbl2 = tblAssignedSamples
                    'tbl3 = tblAssignedSamplesHelper
                    'tbl4 = tblAnalytesHome

                    'setup tables
                    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    intAnalyteID = tbl4.Rows.Item(Count1 - 1).Item("ANALYTEID")

                    If boolUseGroups Then
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "

                        strF2 = strF2 & "ANALYTEID = " & intAnalyteID & " AND "
                        strF2 = strF2 & "INTGROUP = " & intGroup ' & " AND "

                    Else
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                        strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "
                        'strF2 = strF2 & "CHARANALYTE = '" & CleanText(cstr(var3)) & "'" ' And ""
                    End If

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC, ASSAYLEVEL ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug
                    tblNC3 = dv2.ToTable("nc")


                    '****
                    Dim dvNC2 As System.Data.DataView = New DataView(tbl2) 'use All because previous data filtered for accepted analytical runs
                    dvNC2.RowFilter = strF2
                    'strS = "ANALYTEID ASC, MASTERASSAYID ASC, ANALYTEINDEX ASC, LEVELNUMBER ASC"
                    strS = "ANALYTEID ASC, MASTERASSAYID ASC, ANALYTEINDEX ASC, ASSAYLEVEL ASC"
                    strS = "ASSAYLEVEL ASC" 'Amazing, if I use previous sort statement, I get duplicate results in tblNC2
                    strS = "NOMCONC ASC"
                    dvNC2.Sort = strS
                    'now remove assayid and make distinct
                    'Dim tblNC2 As System.Data.DataTable = dvNC2.ToTable("b", True, "ANALYTEID", "MASTERASSAYID", "ANALYTEINDEX", "NOMCONC", "STUDYID", "ASSAYID")
                    Dim tblNC2 As System.Data.DataTable = dvNC2.ToTable("b", True, "ANALYTEID", "MASTERASSAYID", "ANALYTEINDEX", "NOMCONC", "STUDYID", "ASSAYID", "RUNID")
                    Dim tblNC As System.Data.DataTable = dvNC2.ToTable("bnc", True, "NOMCONC")

                    var1 = tblNC.Rows.Count 'debug
                    var2 = tblNC2.Rows.Count 'debug

                    tblNC2.Columns.Add("ASSAYLEVEL", Type.GetType("System.Decimal"))

                    'must do same as Assigned QCs and re-assign levels
                    For Count2 = 0 To tblNC.Rows.Count - 1
                        var1 = NZ(tblNC.Rows(Count2).Item("NOMCONC"), 0)
                        Dim rowsNCB() As DataRow
                        Dim strFNCB As String
                        strFNCB = "NOMCONC = " & var1
                        rowsNCB = tblNC2.Select(strFNCB, "NOMCONC ASC")
                        For Count3 = 0 To rowsNCB.Length - 1
                            rowsNCB(Count3).BeginEdit()
                            rowsNCB(Count3).Item("ASSAYLEVEL") = Count2 + 1
                            rowsNCB(Count3).EndEdit()
                        Next
                    Next

                    'now do the same with tblnc3
                    For Count2 = 0 To tblNC.Rows.Count - 1
                        var1 = NZ(tblNC.Rows(Count2).Item("NOMCONC"), 0)
                        Dim rowsNCB() As DataRow
                        Dim strFNCB As String
                        strFNCB = "NOMCONC = " & var1
                        rowsNCB = tblNC3.Select(strFNCB, "NOMCONC ASC")
                        For Count3 = 0 To rowsNCB.Length - 1
                            rowsNCB(Count3).BeginEdit()
                            rowsNCB(Count3).Item("ASSAYLEVEL") = Count2 + 1
                            rowsNCB(Count3).EndEdit()
                        Next
                    Next

                    '****
                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count


                    'find number of table rows to generate
                    intRowsX = 0
                    ReDim arr1(intNumRuns - 1)
                    For Count2 = 0 To intNumRuns - 1
                        var1 = tblNumRuns.Rows.Item(Count2).Item("RUNID")
                        arr1(Count2) = var1
                    Next


                    var1 = tblNC3.Rows.Count 'debug
                    Call SRSummaryOfBCSC_UseGroups_3(wd, 1, arr1, Count1, intTableID, tblNC2, tblNC3, idTR)

                End If

end1:

                If boolJustTable Then

                    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "") ' NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")

                    If Len(str1) = 0 Then
                    Else
                        strA = strAnal
                        str1 = strA
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                    End If

                    boolJustTable = False
                End If

next1:

            Next
end2:

            'If boolJustTable Then

            '    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")

            '    If Len(str1) = 0 Then
            '    Else
            '        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")
            '    End If

            'End If


        End With
    End Sub


    Sub AnalRefStandards(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strFind As String)

        Dim var1, var2, var3, var4, var5, var6
        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim Count3 As Integer
        Dim Count4 As Integer
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Integer
        Dim dv As System.Data.DataView
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim dr1() As DataRow
        Dim dr2() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        'Dim dg As DataGrid
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim intCols As Short
        Dim wrdselection As Microsoft.Office.Interop.Word.Selection
        Dim strF As String
        Dim numLM, numRM, numRT, numSec, numSec1, numSec2
        Dim boolNew As Boolean
        'Dim ts1 As DataGridTableStyle
        Dim ts2 As DataGridTableStyle
        Dim gs As DataGridColumnStyle
        Dim intTables As Short
        Dim intWTables As Short
        Dim strFOrig As String
        Dim strSOrig As String
        Dim dvW As System.Data.DataView
        Dim strS As String
        Dim col As DataColumn
        Dim row As DataRow
        Dim intRowsTable As Short
        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim tblI As System.Data.DataTable
        Dim rowsI() As DataRow
        Dim intRowsI As Short
        Dim mySel As Microsoft.Office.Interop.Word.Selection

        Dim boolOrigL As Boolean
        Dim boolOrigP As Boolean
        Dim numV As Short
        Dim numM As Short
        Dim strOrig As String
        Dim boolIsIntStd As Boolean = False

        Dim rng1 As Microsoft.Office.Interop.Word.Range

        rng1 = wd.Selection.Range

        Call PositionProgress()

        strOrig = frmH.lblProgress.Text
        numV = frmH.pb1.Value
        numM = frmH.pb1.Maximum
        boolOrigP = frmH.pb1.Visible
        'boolOrigL = frmH.lblProgress.Visible
        boolOrigL = frmH.panProgress.Visible
        frmH.lblProgress.Visible = True
        frmH.pb1.Visible = True
        frmH.pb1.Value = 0

        frmH.panProgress.Visible = True
        frmH.panProgress.Refresh()

        If StrComp(strFind, "None", CompareMethod.Text) = 0 Then
        Else
            wd.Selection.WholeStory()
            'rng1.Select()
            mySel = wd.Selection

            'must first do a find with selection because preceding find was with a range;
            'therefore, [ANALREFTABLE] is not hilited. Must hilite it
            With mySel.Find
                .ClearFormatting()
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                .Execute(FindText:=strFind)
            End With
        End If

        dgv = frmH.dgvCompanyAnalRef
        dv = dgv.DataSource
        tblI = dv.ToTable
        strF = "BOOLINCLUDE = TRUE"
        rowsI = tblI.Select(strF)
        intRowsI = rowsI.Length

        'enter Anal Ref Tables
        numLM = wd.Selection.PageSetup.LeftMargin
        numRM = wd.Selection.PageSetup.RightMargin
        numRT = (8.5 * 72) - numLM - numRM
        str1 = "LM: " & numLM & ", RM: " & numRM & ", RT: " & numRT ''debugging purposes

        '''''''''''wdd.visible = True
        '''''''''''wdd.visible = False

        With wd

            '''''''''''wdd.visible = True

            'add Temp1 bookmark
            wrdselection = wd.Selection()
            With .ActiveDocument.Bookmarks
                .Add(Range:=wrdselection.Range, Name:="Temp1")
                .ShowHidden = False
            End With

            tblA = tblAnalRefStandards
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1"
            strS = "BOOLISINTSTD DESC, CHARCOLUMNNAME ASC"
            rowsA = tblA.Select(strF, strS)
            intRowsTable = intRowsI + 2
            intTables = rowsA.Length
            frmH.pb1.Maximum = intTables

            'start entering data in word tables
            For Count1 = 0 To intTables - 1

                var1 = NZ(rowsA(Count1).Item("BOOLISINTSTD"), 0)
                'If StrComp(var1, "Yes", CompareMethod.Text) = 0 Then
                '    boolIsIntStd = True
                'Else
                '    boolIsIntStd = False
                'End If
                If var1 = 0 Then
                    boolIsIntStd = False
                Else
                    boolIsIntStd = True
                End If

                str1 = rowsA(Count1).Item("CHARANALYTEPARENT")
                If UseAnalyte(str1) Or boolIsIntStd Then

                    frmH.pb1.Value = Count1 + 1
                    frmH.lblProgress.Text = "Generating Analytical Reference Standard Table " & Count1 + 1 & " of " & intTables & "  for " & rowsA(Count1).Item("CHARANALYTEPARENT")
                    frmH.pb1.Refresh()
                    frmH.lblProgress.Refresh()

                    ''wd.visible = True'debugging

                    wrdselection = wd.Selection()
                    .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=intRowsTable, NumColumns:=2, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                    .Selection.Tables.Item(1).Select()
                    Call GlobalTableParaFormat(wd)

                    .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                    With .Selection 'remove initial borders
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                        .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    End With

                    'format paragraph to keep lines with next
                    With .Selection.ParagraphFormat
                        .WidowControl = True
                        .KeepWithNext = True
                        .KeepTogether = True
                    End With

                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    'enter first item

                    Count3 = -1
                    Count4 = 0
                    Try
                        For Count2 = 1 To intRowsTable
                            If Count2 = 1 Then
                                .Selection.Tables.Item(1).Cell(Count2, 2).Select()
                                str2 = rowsA(Count1).Item("CHARANALYTEPARENT")
                                '20181108 LEE: Need to report USERANALYTE or USERIS
                                str2 = GetUserAnalyteNameNoGroup(str2)
                                '
                                str2 = "Compound " & str2 ' rowsA(Count1).Item("CHARANALYTEPARENT")
                                .Selection.Text = str2
                                If boolDoFormulas Then
                                    'unneeded
                                    'Call ChemFormula(wd.Selection.Range, wd) 'call this to sub/superscript stuff
                                End If
                            ElseIf Count2 = 2 Then 'skip this row
                            Else
                                Count3 = Count3 + 1
                                str1 = rowsA(Count1).Item("CHARCOLUMNNAME")
                                var1 = rowsI(Count3).Item("Item")
                                'add : to var1 if not present
                                str2 = Mid(var1, Len(var1), 1)
                                If StrComp(str2, ":", CompareMethod.Text) = 0 Then
                                Else
                                    var1 = var1 & ":"
                                End If
                                var2 = NZ(rowsI(Count3).Item(str1), "")

                                .Selection.Tables.Item(1).Cell(Count2, 1).Select()
                                .Selection.Text = var1
                                .Selection.Tables.Item(1).Cell(Count2, 2).Select()
                                .Selection.Text = var2
                            End If
                            'End If

                        Next
                    Catch ex As Exception

                    End Try


                    'format top row
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                    Try
                        .Selection.Cells.Merge()
                    Catch ex As Exception

                    End Try
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                    '.Selection.Font.Bold = True
                    '.Selection.Font.Size = 13
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .Selection.Tables.Item(1).Cell(intRowsTable, 1).Select()

                    '''''''''''wdd.visible = True

                    '.Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                    '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)


                    '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                    Call MoveOneCellDown(wd)

                    .Selection.TypeParagraph()
                    .Selection.TypeParagraph()

                    var1 = var1 'debug



                    '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")

                    '20140303 Gubbs: Removed this because was getting extra line
                    ' .Selection.TypeParagraph()

                    '.Selection.TypeParagraph()
                    '''''''''''wdd.visible = True
                    ''''''''''''''wdd.visible = False

                Else
                    Dim varaaa
                    varaaa = 1
                End If
            Next

            'add Temp2 bookmark
            wrdselection = wd.Selection()
            With .ActiveDocument.Bookmarks
                .Add(Range:=wrdselection.Range, Name:="Temp2")
                .ShowHidden = False
            End With

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            'var2 = .Selection.Bookmarks.item("Temp2").Start
            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp1")
            'var1 = .Selection.Bookmarks.item("Temp1").Start
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=var2 - var1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)


            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

            frmH.lblProgress.Text = strOrig
            frmH.pb1.Value = 0
            frmH.pb1.Maximum = numM
            frmH.pb1.Value = numV
            'frmH.lblProgress.Visible = boolOrigL
            frmH.panProgress.Visible = boolOrigL
            'frmH.pb1.Visible = boolOrigP
            frmH.lblProgress.Refresh()
            frmH.pb1.Refresh()
            frmH.panProgress.Refresh()

        End With

    End Sub

    Sub SummaryTableMulti(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim var1, var2, var3, var4, var5, var6
        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim Count3 As Integer
        Dim Count4 As Integer
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Integer
        Dim dv As system.data.dataview
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim tbl2 As System.Data.DataTable
        Dim dr1() As DataRow
        Dim dr2() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim intCols As Short
        Dim wrdselection As Microsoft.Office.Interop.Word.selection
        Dim strF As String
        Dim strS As String
        Dim numLM, numRM, numRT, numSec1, numSec2
        Dim boolNew As Boolean
        Dim strFOrig As String
        Dim strAnal As String
        Dim intRow As Short

        'enter Contributing Personnel Table
        frmH.lblProgress.Text = "Enter Summary Table..."
        frmH.lblProgress.Refresh()
        ctPB = ctPB + 1
        If ctPB > ctPBMax Then
            ctPB = 1
        End If
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()

        '''''wdd.visible = True

        With wd

            'determine number of rows
            dgv = frmH.dgvSummaryData
            dv = dgv.DataSource
            Dim tbl3 As System.Data.DataTable = dv.ToTable()
            Dim rows3() As DataRow
            strF = "boolInclude = -1"
            strS = "INTORDER ASC"
            rows3 = tbl3.Select(strF, strS)
            intRows = rows3.Length

            'find margins
            numLM = wd.Selection.PageSetup.LeftMargin
            numRM = wd.Selection.PageSetup.RightMargin
            numRT = (8.5 * 72) - numLM - numRM

            ''generate table
            'wrdselection = wd.selection()


            'determine page orientation
            Dim tblAF As System.Data.DataTable
            Dim rowsAF() As DataRow
            Dim intRowsAF As Short

            tblAF = tblAppFigs
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGAPPFIGS = 7"
            rowsAF = tblAF.Select(strF)

            intRowsAF = rowsAF.Length

            tbl1 = tblMethodValData


            For Count2 = 1 To ctAnalytes

                'insert a page break
                If Count2 = 1 Then
                Else
                    .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                End If

                If intRowsAF = 0 Then
                    Call PageSetup(wd, "P")
                Else
                    Call PageSetup(wd, NZ(rowsAF(0).Item("CHARPAGEORIENTATION"), "P"))
                End If

                '.ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=intRows, NumColumns:= _
                '2, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=var1)

                'generate table
                wrdselection = wd.Selection()
                .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=intRows, NumColumns:= _
                               2, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .Selection.Tables.Item(1).Select()
                '20180829 LEE: Don't set globaltable, just let be normal
                'Call GlobalTableParaFormat(wd)

                .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False

                With .Selection 'remove initial borders
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                End With
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                strAnal = arrAnalytes(1, Count2)

                'begin entering data
                Try
                    For Count1 = 0 To intRows - 1
                        var1 = NZ(rows3(Count1).Item("charRowName"), "")

                        .Selection.Tables.Item(1).Cell(Count1 + 1, 1).Select()
                        .Selection.TypeText(Text:=var1)
                        intRow = FindSumRow(var1, tbl1)
                        If intRow = -1 Then
                            var2 = NZ(rows3(Count1).Item("charValue"), "")
                            Select Case var1

                                '20170821 LEE: don't strip analyte
                                Case "Analytes"
                                    var2 = strAnal 'FindCStuff(var2, strAnal)
                                Case "Internal Standard"
                                    'var2 = FindCStuff(var2, strAnal)
                                Case "Standard Curve Dynamic Range"
                                    'var2 = FindCStuff(var2, strAnal)
                                Case "Regression Type"
                                    'var2 = FindCStuff(var2, strAnal)
                                Case "Validation References"
                                    'var2 = FindCStuff(var2, strAnal)
                                Case Else

                            End Select

                        Else
                            var2 = tbl1.Rows(intRow).Item(strAnal)
                        End If
                        .Selection.Tables.Item(1).Cell(Count1 + 1, 2).Select()
                        .Selection.TypeText(Text:=NZ(var2, "[None]"))
                    Next
                Catch ex As Exception
                    var3 = ex.Message
                    var3 = var3
                End Try
                

                'enter title
                .Selection.Tables.Item(1).Cell(1, 1).Select()
                .Selection.InsertRowsAbove(1)
                .Selection.Tables.Item(1).Cell(1, 1).Select()
                .Selection.SelectRow()
                .Selection.Cells.Merge()
                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                '.Selection.Font.Size = .Selection.Font.Size + 2
                .Selection.Font.Bold = True
                .Selection.ParagraphFormat.LineSpacing = 24 'LinesToPoints(2)
                .Selection.TypeText(Text:=strAnal & " Method Validation Summary")

                .Selection.Rows.HeadingFormat = True
                .Selection.Tables(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                .Selection.Tables(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

                Call MoveOneCellDown(wd)

                ''''''''wdd.visible = True

            Next

        End With

    End Sub

    Sub SummaryTableOne(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim var1, var2, var3, var4, var5, var6
        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim Count3 As Integer
        Dim Count4 As Integer
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Integer
        Dim dv As system.data.dataview
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim dr1() As DataRow
        Dim dr2() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim intCols As Short
        Dim wrdselection As Microsoft.Office.Interop.Word.selection
        Dim strF As String
        Dim strS As String
        Dim numLM, numRM, numRT, numSec1, numSec2
        Dim boolNew As Boolean
        Dim strFOrig As String

        'enter Contributing Personnel Table
        frmH.lblProgress.Text = "Enter Summary Table..."
        frmH.lblProgress.Refresh()
        ctPB = ctPB + 1
        If ctPB > ctPBMax Then
            ctPB = 1
        End If
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        With wd

            'determine number of rows
            dgv = frmH.dgvSummaryData
            dv = dgv.DataSource
            Dim tbl3 As System.Data.DataTable = dv.ToTable()
            Dim rows3() As DataRow
            strF = "boolInclude = -1"
            strS = "INTORDER ASC"
            rows3 = tbl3.Select(strF, strS)
            intRows = rows3.Length

            'find margins
            numLM = wd.Selection.PageSetup.LeftMargin
            numRM = wd.Selection.PageSetup.RightMargin
            numRT = (8.5 * 72) - numLM - numRM

            'generate table
            wrdselection = wd.Selection()

            'insert a page break'DO THIS IN EARLIER CODE!
            '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

            'determine page orientation
            Dim tblAF As System.Data.DataTable
            Dim rowsAF() As DataRow
            Dim intRowsAF As Short

            tblAF = tblAppFigs
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGAPPFIGS = 7"
            rowsAF = tblAF.Select(strF)

            intRowsAF = rowsAF.Length


            If intRowsAF = 0 Then
                Call PageSetup(wd, "P")
            Else
                Call PageSetup(wd, NZ(rowsAF(0).Item("CHARPAGEORIENTATION"), "P"))
            End If

            .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=intRows, NumColumns:= _
            2, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
            .Selection.Tables.Item(1).Select()
            '20180829 LEE: Don't set globaltable, just let be normal
            'Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False

            With .Selection 'remove initial borders
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                '.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            End With
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            'begin entering data
            For Count1 = 0 To intRows - 1
                var1 = NZ(rows3(Count1).Item("charRowName"), "")
                var2 = NZ(rows3(Count1).Item("charValue"), "")
                .Selection.Tables.Item(1).Cell(Count1 + 1, 1).Select()
                .Selection.TypeText(Text:=var1)
                .Selection.Tables.Item(1).Cell(Count1 + 1, 2).Select()
                .Selection.TypeText(Text:=NZ(var2, "[None]"))
            Next

            'enter title
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.InsertRowsAbove(1)
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.SelectRow()
            .Selection.Cells.Merge()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            '.Selection.Font.Size = .Selection.Font.Size + 2
            .Selection.Font.Bold = True
            .Selection.ParagraphFormat.LineSpacing = 24 'LinesToPoints(2)
            .Selection.TypeText(Text:="Method Validation Summary")

            .Selection.Rows.HeadingFormat = True
            .Selection.Tables(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            .Selection.Tables(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

            Call MoveOneCellDown(wd)

        End With

    End Sub

    Sub SummaryTableAppendix(ByRef wd As Microsoft.Office.Interop.Word.Application)

        If frmH.rbMultValYes.Checked Then
            Call SummaryTableMulti(wd)
        Else
            Call SummaryTableOne(wd)
        End If

    End Sub

    Sub ContributingPersonnel(ByRef wd As Microsoft.Office.Interop.Word.Application, ByVal boolTitle As Boolean)

        Dim var1, var2, var3, var4, var5, var6
        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim Count3 As Integer
        Dim Count4 As Integer
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Integer
        Dim dv As System.Data.DataView
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim dr1() As DataRow
        Dim dr2() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim intCols As Short
        Dim wrdselection As Microsoft.Office.Interop.Word.Selection
        Dim strF As String
        Dim numLM, numRM, numRT, numSec1, numSec2
        Dim boolNew As Boolean
        Dim strFOrig As String

        'enter Contributing Personnel Table
        frmH.lblProgress.Text = "Enter Contributing Personnel Table..."
        frmH.lblProgress.Refresh()
        ctPB = ctPB + 1

        If ctPB > frmH.pb1.Maximum Then
            ctPB = 1
        End If
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()

        Dim strC As String = ""

        With wd

            'determine number of rows
            dgv = frmH.dgvContributingPersonnel
            dv = dgv.DataSource
            strFOrig = dv.RowFilter
            strF = "BOOLINCLUDESIGONTABLEPAGE = -1 AND id_tblStudies = " & id_tblStudies
            dv.RowFilter = strF
            str1 = "intOrder ASC"
            dv.Sort = str1
            intRows = dv.Count

            'find margins
            numLM = wd.selection.pagesetup.leftmargin
            numRM = wd.selection.pagesetup.rightmargin
            numRT = (8.5 * 72) - numLM - numRM

            wrdselection = wd.selection()
            'set right tab at right margin
            .Selection.ParagraphFormat.TabStops.Add(Position:=numRT, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots)
            For Count1 = 0 To intRows - 1
                '.Selection.ParagraphFormat.TabStops.Add(Position:=wd.InchesToPoints(7), _
                '  Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                '.Selection.ParagraphFormat.TabStops.ClearAll()
                '.ActiveDocument.DefaultTabStop = wd.InchesToPoints(0.5)
                '.Selection.ParagraphFormat.TabStops.Add(Position:=wd.InchesToPoints(7), _
                '  Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots)
                var1 = NZ(dv(Count1).Item("charCPTitle"), "")
                var2 = NZ(dv(Count1).Item("charCPName"), "")
                var3 = NZ(dv(Count1).Item("charCPDegree"), "")

                '20180827 LEE:
                'New logic
                'If Len(var3) = 0 Then
                '    var4 = var1 & ChrW(9) & var2
                'Else
                '    var4 = var1 & ChrW(9) & var2 & ", " & var3
                'End If

                If boolTitle Then
                    If Len(var3) = 0 Then
                        var4 = var2 & ChrW(9) & var1
                    Else
                        var4 = var2 & ", " & var3 & ChrW(9) & var1
                    End If
                Else
                    If Len(var3) = 0 Then
                        var4 = var2
                    Else
                        var4 = var2 & ", " & var3
                    End If
                End If

                If Count1 = 0 Then
                    strC = var4
                Else
                    strC = strC & ChrW(10) & var4
                End If

                '.Selection.TypeText(Text:=var4)
                'If Count1 = intRows - 1 Then
                '    .Selection.TypeParagraph()
                'Else
                '    .Selection.TypeParagraph()
                '    .Selection.TypeParagraph()
                'End If
            Next

            .Selection.TypeText(Text:=strC)

            If boolTitle Then
                .Selection.ParagraphFormat.TabStops.ClearAll()
                .ActiveDocument.DefaultTabStop = (0.5 * 72) 'wd.InchesToPoints(0.5)
                '.Selection.TypeParagraph()
            Else

            End If


            'skip this next section for a while
            GoTo skip1

            'now enter signature blocks, if needed
            strF = "boolIncludeSigOnTablePage = " & -1 & " AND id_tblStudies = " & id_tblStudies
            dv = dgv.DataSource
            dv.RowFilter = strF
            str1 = "intOrder ASC"
            dv.Sort = str1
            intRows = dv.Count
            If intRows = 0 Then
            Else
                Select Case intRows
                    Case 1
                        intCols = 2
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow
                    Case Is > 1
                        intCols = 3
                        var1 = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed
                End Select
                wrdselection = wd.selection()
                .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=2, NumColumns:= _
                intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=var1)
                '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
                .selection.Tables.item(1).Select()
                Call GlobalTableParaFormat(wd)

                .Selection.Tables.item(1).Rows.AllowBreakAcrossPages = False

                With .selection 'remove initial borders
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End With
                .selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                numSec1 = numRT * 0.45
                numSec2 = numRT * 0.1

                Select Case intRows
                    Case 1
                        Count3 = 1
                        .selection.Tables.item(1).cell(Count3, 1).select()

                        .Selection.Tables.item(1).Rows.item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                        .Selection.Tables.item(1).Rows.item(Count3).Height = 72 'InchesToPoints(1)
                        .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                        .Selection.Tables.item(1).Rows.item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                        var1 = NZ(dv(0).Item("charCPName"), "")
                        var2 = NZ(dv(0).Item("charCPDegree"), "")
                        var5 = NZ(dv(0).Item("charCPTitle"), "")
                        If Len(var2) = 0 Then
                            var3 = var1 & " / Date" & Chr(10) & var5
                        Else
                            var3 = var1 & ", " & var2 & " / Date" & Chr(10) & var5
                        End If
                        .Selection.TypeText(Text:=var3)
                        .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)

                    Case Is > 1
                        .selection.Columns.item(2).width = numSec2
                        .selection.Columns.item(1).width = numSec1
                        .selection.Columns.item(3).width = numSec1

                        'format columns
                        boolNew = False
                        int1 = CInt(intRows / 2)
                        Count2 = 1
                        Count3 = 1
                        .selection.Tables.item(1).cell(Count3, 1).select()
                        For Count1 = 0 To intRows - 1
                            If Count1 = 0 Or boolNew Then
                                .Selection.Tables.item(1).Rows.item(Count3).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly
                                .Selection.Tables.item(1).Rows.item(Count3).Height = 72 'InchesToPoints(1)
                                '.Selection.Tables.item(1).Rows.AllowBreakAcrossPages = False
                                .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=intCols)
                                .Selection.Tables.item(1).Rows.item(Count3 + 1).HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto
                                '.Selection.Tables.item(1).Rows.Height = InchesToPoints(0)
                                boolNew = False
                            End If
                            '.selection.Tables.item(1).cell(Count3, 1).select()
                            '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                            '.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                            var1 = dv(Count1).Item("charCPName")
                            var2 = NZ(dv(Count1).Item("charCPDegree"), "")
                            var5 = NZ(dv(Count1).Item("charCPTitle"), "")
                            If Len(var2) = 0 Then
                                var3 = var1 & " / Date" & Chr(10) & var5
                            Else
                                var3 = var1 & ", " & var2 & " / Date" & Chr(10) & var5
                            End If
                            .Selection.TypeText(Text:=var3)
                            .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            If Count1 = intRows - 1 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                            ElseIf Count1 = Count2 Then
                                .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=1)
                                Count2 = Count2 + 2
                                Count3 = Count3 + 2
                                boolNew = True
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                .Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                                .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
                            Else
                                .selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=2)
                            End If
                        Next
                End Select
            End If
            .selection.typeparagraph()
        End With

skip1:

        'set dv back to normal. It's affecting dgContributingPersonnel
        dv.RowFilter = strFOrig
        'Call frmH.doCPCancel()



    End Sub

    Sub QATable(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim var1, var2, var3, var4, var5, var6
        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim Count3 As Integer
        Dim Count4 As Integer
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Integer
        Dim dv As System.Data.DataView
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim dr1() As DataRow
        Dim dr2() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim boolInclude As Boolean
        Dim dg As DataGrid
        Dim intRows As Short
        Dim intCols As Short
        Dim arrCols(2, 15)
        '1=col header, 2=mapping name
        Dim ts As DataGridTableStyle
        Dim gs As DataGridColumnStyle
        Dim boolHit As Boolean
        Dim wrdselection As Microsoft.Office.Interop.Word.Selection
        Dim strCurr As String
        Dim strPrev As String
        Dim fontsize

        'enter QA Table
        frmH.lblProgress.Text = "Entering QA Table Table..."
        frmH.lblProgress.Refresh()
        ctPB = ctPB + 1
        'If ctPB > ctPBMax Then
        '    ctPB = 1
        'End If
        'frmH.pb1.Value = ctPB
        'frmH.pb1.Refresh()

        ''''''''''wdd.visible = True
        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size ' wd.Selection.Font.Size

        If ctPB > frmH.pb1.Maximum Then
            ctPB = 1
        End If
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()

        With wd
            dg = frmH.dgQATable
            dv = dg.DataSource
            tbl1 = tblQATableTemp
            'tbl2 = frmH.tblQACriticalPhases
            'str1 = "id_tblQACriticalPhases > 0"
            'dr2 = tbl2.Select(str1, "id_tblQACriticalPhases ASC")
            'ct2 = dr2.Length
            ts = dg.TableStyles(0)
            Count1 = 0
            For Each gs In ts.GridColumnStyles
                Count1 = Count1 + 1
                arrCols(1, Count1) = gs.HeaderText
                arrCols(2, Count1) = gs.MappingName
            Next
            intCols = Count1
            intRows = dv.Count

            wrdselection = wd.Selection()
            .ActiveDocument.Tables.Add(Range:=wrdselection.Range, NumRows:=intRows + 3, NumColumns:= _
                intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
            .Selection.Tables.Item(1).Select()
            .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False

            'optimize cells leftpadding
            .Selection.Tables.Item(1).Select()
            Call GlobalTableParaFormat(wd)

            .Selection.Tables.Item(1).LeftPadding = 5 '2.5
            .Selection.Tables.Item(1).RightPadding = 5

            With .Selection 'remove initial borders
                If BOOLQAEVENTBORDER Then
                Else
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                End If
            End With
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            '.selection.Font.Size = 11

            'this code for first column of QA table
            .Selection.Tables.Item(1).Cell(4, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
            With .Selection.ParagraphFormat
                .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                .LeftIndent = 18 'InchesToPoints(0.25)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
            End With
            With .Selection.ParagraphFormat
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .FirstLineIndent = -18 'InchesToPoints(-0.25)
            End With

            '''''wdd.visible = True

            Count3 = 3
            Count2 = 1
            For Count4 = 0 To intRows - 1 'ct2 '3
                Count3 = Count3 + 1
                .Selection.Tables.Item(1).Cell(Count3, 1).Select()
                If Count4 = 0 Then
                    strCurr = dv(Count4).Item("charUserLabel")
                    strPrev = "GGG"
                Else
                    strPrev = strCurr
                    strCurr = dv(Count4).Item("charUserLabel")
                End If
                If Count4 = 0 Then
                    str1 = Count2 & Chr(9) & strCurr
                ElseIf StrComp(strCurr, strPrev, CompareMethod.Text) = 0 Then
                    str1 = ""
                Else
                    Count2 = Count2 + 1
                    .Selection.InsertRows(1)
                    Count3 = Count3 + 1 'increment one more to allow for spaces between critical phases
                    .Selection.Tables.Item(1).Cell(Count3, 1).Select()
                    str1 = Count2 & Chr(9) & strCurr
                End If
                .Selection.TypeText(str1)
                .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                'enter Laboratory Procedure
                For Count1 = 2 To intCols
                    str1 = arrCols(2, Count1)
                    var1 = NZ(dv(Count4).Item(str1), "")
                    If Len(var1) = 0 Then
                    Else
                        'var1 = Format(CDate(var1), LDateFormat)
                        Try
                            var1 = Format(CDate(var1), LDateFormat)
                        Catch ex As Exception
                            'leave as text
                        End Try
                    End If
                    .Selection.TypeText(Text:=var1)
                    If Count1 = intCols And Count4 = intRows - 1 Then
                    Else
                        .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                    End If
                Next
                ''
                '''debug.writeline(Count4 & ", " & strCurr & ", " & strPrev)
            Next

            'enter table headings
            .Selection.Tables.Item(1).Cell(2, 1).Select()
            For Count2 = 1 To intCols
                'var1 = rng1.Offset(-1, Count2).Value
                var1 = arrCols(1, Count2)
                'replace spaces with carriage returns
                var2 = Replace(var1, " ", Chr(10), 1, -1, CompareMethod.Text)
                .Selection.Tables.Item(1).Cell(2, Count2).Select()
                If Count2 < gSDMax Then
                    If BOOLQAEVENTBORDER Then
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    End If
                End If

                '20171220 LEE: Do not set table size, use the style default table
                '.Selection.Font.Size = fontsize - 1
                .Selection.Font.Bold = True
                .Selection.TypeText(Text:=var2)
            Next

            'format table
            '.selection.Tables.item(1).Select()
            '.Selection.Rows.AllowBreakAcrossPages = False


            'With .selection.PageSetup
            '    var1 = .LeftMargin
            '    var2 = .RightMargin
            'End With
            Dim numTot, numSec, num1

            '' 'must make doc visible or column settings just don't work

            .Selection.Tables.Item(1).Columns.Item(1).PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
            '.selection.Tables.item(1).Columns.item(1).PreferredWidth = 144 'Comments
            '.selection.Tables.item(1).Columns.item(1).select()
            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
            '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            '.Selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            '.selection.Tables.item(1).Columns.item(1).autofit()
            '.selection.Tables.item(1).Columns.item(1).Select()

            num1 = .Selection.Tables.Item(1).Columns.Item(1).Width '166
            .Selection.Tables.Item(1).Columns.Item(1).Width = num1 * 2

            'numTot = (8.5 * 72) - num1 - var1 - var2
            'numSec = numTot / (intCols - 1)
            '.selection.Tables.item(1).Select()

            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)

            '.selection.Tables.item(1).Columns.item(1).Select()

            'For Count1 = 2 To intCols
            '    .selection.Move(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Count:=1)
            '    .selection.SelectColumn()
            '    '.selection.Columns.PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPoints
            '    .selection.Columns.width = numSec 'InchesToPoints(1.2)

            'Next

            'enter and merge top heading
            If .Selection.Tables.Item(1).Columns.Count < 4 Then
                '.selection.Tables.item(1).Cell(1, 2).Select()
            Else
                .Selection.Tables.Item(1).Cell(1, 3).Select()
                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                Try
                    .Selection.Cells.Merge()
                Catch ex As Exception

                End Try
                .Selection.Font.Bold = True
                .Selection.TypeText(Text:="Date Findings Submitted To:")
                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
            End If

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.InsertRows(1)
            Try
                .Selection.Cells.Merge()
            Catch ex As Exception

            End Try
            .Selection.Font.Bold = True
            'fonts = .Selection.Font.Size
            '.Selection.Font.Size = 12
            .Selection.TypeText(Text:="QA INSPECTION DATES")
            .Selection.TypeParagraph()
            '.Selection.Font.Size = fonts


            ''''''''''wdd.visible = True

            .Selection.Tables.Item(1).Select()
            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
            '.selection.Tables.item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
            autofitWindow(wd, 2)
            .Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)

            'Dim intr As Short
            'Dim intc As Short
            'intr = .selection.Tables.item(1).rows.count
            'intc = .selection.Tables.item(1).columns.count
            '.selection.Tables.item(1).cell(intr, intc).select()
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdWord, Count:=1, Extend:=True)
            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=3)
            '.selection.typeparagraph()
            '***end QA Table

            ''''''''''''''''wdd.visible = False

        End With

        'refresh form
        'frmH.Refresh()

    End Sub

    Sub SRSummaryOfLSR_2(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64, ByVal boolSA As Boolean)

        'boolSA: Sample Assignement: True if samples have been assigned, False if no samples have been assigned

        Dim boolNA As Boolean = False
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
        Dim arrAnalyticalRuns(1, 1)
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
        Dim arrTemp(2, 50)
        Dim num1 As Object
        Dim num2 As Object
        Dim num3 As Object
        Dim arrBCStdActual(1, 1)

        Dim ctLegend As Short
        Dim lng1 As Integer
        Dim lng2 As Integer
        Dim boolPortrait As Boolean
        Dim intLastAnal As Short
        Dim ctCols 'number of columns in a table
        Dim strSub1 As String
        Dim strSub2 As String
        Dim pos1 As Short
        Dim pos2 As Short
        Dim dvDo As System.Data.DataView
        Dim intDo As Short
        Dim strDo As String
        Dim bool As Boolean
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim drows() As DataRow
        Dim arrRegCon(4, 10)
        Dim arrDt(10)
        Dim inttemprows As Short
        Dim strTName As String
        Dim intRP As Short 'number of regression parameters
        Dim strRegressionType As String 'regression type
        Dim strWeighting As String
        Dim intTRows As Short
        Dim strTempInfo As String
        Dim strA As String
        Dim strF As String

        Dim rowsARS() As DataRow
        Dim strFARS As String
        Dim intARS As Short

        Dim arrRunID(1)
        Dim intRunID As Int64
        Dim boolSRegr As Boolean = True 'Single Regression
        Dim intNumRegr As Short
        Dim arrRegrType(2, 1) 'need for legend 1=Type, 2=Wt
        Dim intRegrCt As Short
        Dim intNumRegrWt As Short

        Dim strRegrLegend As String
        Dim intLeg As Short
        Dim intLegStart As Short

        Dim intRows As Short = 0
        Dim intRow As Short = 0

        Dim strMatrix As String
        Dim strTNameO As String
        Dim strAnalyteDescription As String

        Dim boolAreaRatio As Boolean = True 'if regressions use area ratio or analyte peak area

        var1 = arrAnalytes(1, 1)

        ''''''''wdd.visible = True

        'dvDo = frmH.dgvReportTableConfiguration.DataSource
        'strTName = "Summary of Regression Constants"
        'intDo = FindRowDVByCol(strTName, dvDo, "Table")

        Dim intTableID As Short
        intTableID = 2

        Dim strWRunId As String = GetWatsonColH(intTableID)

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

        Dim fonts
        Dim fontsize
        'fontsize = wd.Selection.Font.Size
        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        Dim charFCID As String
        Dim rowsFC() As DataRow = tblReportTable.Select("ID_TBLREPORTTABLE = " & idTR)
        charFCID = NZ(rowsFC(0).Item("CHARFCID"), "NA")

        Dim strM As String
        Dim strM1 As String
        Dim strAnal As String
        Dim strAnalC As String
        Dim boolExPSAE As Boolean = False

        Dim tblAS As DataTable = tblAssignedSamples '20190304 LEE:

        Dim intAnalyteID As Int64

        If frmH.chkPSAE.Checked Then
            boolExPSAE = True
        End If

        '
        '20190304 LEE:
        'check if samples are to be assigned
        str1 = "ID_TBLCONFIGREPORTTABLES" 'column name
        'int2 = FindRowDVNumByCol(int1, dv, str1)

        int2 = intRow '

        With wd

            For Count1 = 1 To ctAnalytes

                Dim boolJustTable As Boolean = False

                strTName = strTNameO

                ctLegend = 0

                var1 = tblAnalyteGroups.Rows(Count1 - 1).Item("ANALYTEID")
                If Count1 = 17 Then 'debug
                    var1 = var1
                End If
                If boolUseGroups Then
                    intGroups = tblAnalyteGroups.Rows(Count1 - 1).Item("INTGROUP")
                    strMatrix = tblAnalyteGroups.Rows(Count1 - 1).Item("MATRIX")
                    strAnalC = tblAnalyteGroups.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                    gstrAnal = tblAnalyteGroups.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C") ' arrAnalytes(1, Count1)
                    strAnal = tblAnalyteGroups.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION") ' arrAnalytes(1, Count1)
                    intAnalyteID = tblAnalyteGroups.Rows(Count1 - 1).Item("ANALYTEID")
                Else
                    strAnalC = arrAnalytes(14, Count1)
                    gstrAnal = arrAnalytes(1, Count1)
                    strAnal = arrAnalytes(1, Count1)
                End If

                gnumAnal = Count1

                Dim intGroup As Short = tblAnalyteGroups.Rows(Count1 - 1).Item("INTGROUP")

                'check if table is to be generated
                strDo = arrAnalytes(1, Count1) 'record column name

                If UseAnalyte(CStr(strDo)) Then
                Else
                    GoTo next1
                End If

                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                If bool Then 'continue

                    Dim arrLegend(4, 20)

                    intTCur = intTCur + 1

                    'check if there are any runs for this group
                    Dim rowIG() As DataRow
                    Dim intIG As Short
                    If boolSA Then
                        'check for information in tblAS
                        strF = "ID_TBLREPORTTABLE = " & idTR & " AND INTGROUP = " & intGroups
                        rowIG = tblAS.Select(strF)
                    Else
                        If boolExPSAE Then
                            strF = "INTGROUP = " & intGroups & " AND RUNTYPEID > 0"
                        Else
                            strF = "INTGROUP = " & intGroups & " AND RUNTYPEID <> 3"
                        End If
                        rowIG = tblCalStdGroupAssayIDsAcc.Select(strF)
                    End If
                    intIG = rowIG.Length
                    If intIG = 0 Then
                        '20190304 LEE: Make a table placeholder
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
                    End If

                    intLeg = 0
                    intLegStart = 96

                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    strM = "Creating " & strTName & " For " & arrAnalytes(1, Count1) & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    int2 = arrAnalytes(3, Count1) 'analyteindex

                    'record strregressiontype
                    int1 = FindRow("Regression", tblWatsonAnalRefTable, "Item")
                    int2 = FindRow("Weighting", tblWatsonAnalRefTable, "Item")
                    strRegressionType = NZ(tblWatsonAnalRefTable.Rows(int1).Item(Count1), "NA")
                    strWeighting = NZ(tblWatsonAnalRefTable.Rows(int2).Item(Count1), "NA")

                    'find number of regression parameters
                    'Dim dvT as system.data.dataview = New DataView(tblRegCon)
                    'use tblRegCon to get parameters
                    Dim dvT1 As System.Data.DataView

                    If boolSA Then
                        dvT1 = New DataView(tblRegConAll)
                    Else
                        If boolExPSAE Then
                            dvT1 = New DataView(tblRegConAll)
                        Else
                            dvT1 = New DataView(tblRegCon)
                        End If
                    End If


                    'rbAnalRunsBioanalysis

                    If boolSA Then '20190304 LEE:
                        'get unique runid from rowig
                        strF = "ID_TBLREPORTTABLE = " & idTR & " AND INTGROUP = " & intGroups
                        rowIG = tblAS.Select(strF)
                        Dim dvG As DataView = New DataView(tblAS, strF, "RUNID ASC", DataViewRowState.CurrentRows)
                        Dim tblG As System.Data.DataTable = dvG.ToTable("a", True, "RUNID")

                        For Count2 = 0 To tblG.Rows.Count - 1
                            If Count2 = 0 Then
                                str2 = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND (RUNID = " & tblG.Rows(Count2).Item("RUNID")
                            Else
                                str2 = str2 & " OR RUNID = " & tblG.Rows(Count2).Item("RUNID")
                            End If
                        Next
                        str1 = str2 & ")"
                    Else
                        If frmH.chkAll.Checked Then
                            str1 = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & strMatrix & "' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        ElseIf frmH.chkPSAE.Checked = False Then 'exlude PSAE
                            str1 = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND RUNTYPEID <> 3 AND SAMPLETYPEID = '" & strMatrix & "' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        ElseIf frmH.chkPSAE.Checked Then 'Include PSAE Accepted
                            str1 = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND RUNTYPEID > 0 AND SAMPLETYPEID = '" & strMatrix & "' AND RUNANALYTEREGRESSIONSTATUS = 3"
                        End If
                    End If


                    dvT1.RowFilter = str1

                    If BOOLINCLUDEDATE Then
                        dvT1.Sort = "RUNSTARTDATE ASC, RUNID ASC"
                    Else
                        dvT1.Sort = "RUNID ASC"
                    End If

                    '******

                    Dim strFFF As String
                    Dim tblRID As DataTable = dvT1.ToTable
                    strFFF = GetARSRuns(tblRID, intAnalyteID, "", boolSA)

                    If Len(strFFF) = 0 Then
                    Else
                        strFFF = "(" & strFFF & ")"
                        str1 = str1 & " AND " & strFFF
                    End If

                    '******

                    dvT1.RowFilter = str1


                    ''find number of table rows
                    '****
                    Dim tblRunID As System.Data.DataTable
                    intTRows = 0
                    tblRunID = dvT1.ToTable("b", True, "RUNID")
                    intTRows = tblRunID.Rows.Count

                    Dim tblOQ As System.Data.DataTable = dvT1.ToTable("c", True, "RUNID", "NM", "VEC")
                    Dim rowsOQ() As DataRow


                    '****

                    Dim tblT As System.Data.DataTable = dvT1.ToTable("a", True, "REGRESSIONPARAMETERID")
                    intRP = tblT.Rows.Count

                    'determine if there is more than one regr type
                    'Dim tblSR As System.Data.DataTable = dvT1.ToTable("sr", True, "REGRESSIONTEXT")
                    '20151218 LEE: Need to evaluate different weightings as well

                    Dim tblSR As System.Data.DataTable = dvT1.ToTable("sr", True, "REGRESSIONTEXT", "WEIGHTINGFACTOR") 'this comes from tblRegCon
                    intNumRegrWt = tblSR.Rows.Count

                    Dim tblSRegr As System.Data.DataTable = dvT1.ToTable("sr", True, "REGRESSIONTEXT") 'this comes from tblRegCon
                    'intNumRegr = tblSR.Rows.Count
                    intNumRegr = tblSRegr.Rows.Count

                    'If intNumRegr = 1 Then
                    '    var1 = tblSR.Rows(0).Item("WEIGHTINGFACTOR")
                    '    strWeighting = GetWt(var1)
                    'End If

                    If intNumRegrWt = 1 Then
                        var1 = tblSR.Rows(0).Item("WEIGHTINGFACTOR")
                        strWeighting = GetWt(var1)
                    End If

                    'now set dvT to regconall
                    Dim dvT As System.Data.DataView

                    If boolExPSAE Then
                        dvT = New DataView(tblRegConAll)
                    Else
                        dvT = New DataView(tblRegCon)
                    End If

                    'ReDim arrRegrType(2, intNumRegr)
                    'If intNumRegr = 1 Then
                    '    boolSRegr = True
                    'Else
                    '    boolSRegr = False
                    'End If
                    'For Count2 = 1 To intNumRegr
                    '    arrRegrType(1, Count2) = tblSR.Rows(Count2 - 1).Item("REGRESSIONTEXT")
                    '    'need for legend 1=Type, 2=Wt
                    'Next

                    ReDim arrRegrType(2, intNumRegrWt)
                    'if intNumRegrWt > 1, then need to add an additional column for Weighting
                    If intNumRegrWt = 1 Then
                        boolSRegr = True
                    Else
                        boolSRegr = False
                    End If
                    For Count2 = 1 To intNumRegr 'keep this as tblSR
                        arrRegrType(1, Count2) = tblSR.Rows(Count2 - 1).Item("REGRESSIONTEXT")
                        'need for legend 1=Type, 2=Wt
                    Next

                    inttemprows = intTRows
                    intRegrCt = intTRows
                    ReDim arrRegCon(intRP + 3, intRegrCt)
                    '1=RunID, 2=RegrParameter1, etc
                    ReDim arrRunID(intRegrCt)
                    ReDim arrDt(intRegrCt)

                    Dim intCols As Short
                    If boolSRegr Then
                        intCols = intRP + 2
                    Else
                        intCols = intRP + 2 + 2
                    End If

                    If BOOLREGRULOQ Then
                        intCols = intCols + 2
                    End If

                    Count2 = 0
                    '1=RUNID, 2=AnalyteIndex, 3=REGRESSIONPARAMETERID(1=Slope, 2=YInt, 3=R2),4=PARAMETERVALUE
                    '1=RUNID,  2=Slope, 3=YInt, 4=R2
                    '1=RUNID, intRP parameters, intRP+1=R2
                    'get info from tblregcon
                    'RUNANALYTEREGRESSIONSTATUS
                    'this should not take into account rbAnalRunsShowAll
                    'it should only take into account PSAE or Accepted

                    'If frmH.rbAnalRunsShowAll.Checked Then
                    '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0"
                    'ElseIf frmH.rbAnalRunsExclPSAE.Checked Then
                    '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3"
                    'Else 'Bioanalysis
                    '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0"
                    'End If
                    'If boolIncludePSAE Then
                    '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0"
                    'Else
                    '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3"
                    'End If
                    'str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                    'drows = tblRegCon.Select(str1)

                    If boolUseGroups Then
                        'If frmH.rbAnalRunsShowAll.Checked Then
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID > 0 AND RUNID = " & var1
                        'ElseIf frmH.rbAnalRunsExclPSAE.Checked Then
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID <> 3 AND RUNID = " & var1
                        'Else 'Accepted
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"

                        'End If
                        If frmH.chkAll.Checked Then
                            strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        ElseIf frmH.chkPSAE.Checked = False Then 'Excludue PSAE
                            strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID <> 3 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        Else 'Accepted including PSAE
                            strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        End If
                    Else
                        'If frmH.rbAnalRunsShowAll.Checked Then
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        'ElseIf frmH.rbAnalRunsExclPSAE.Checked Then
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        'Else 'Accepted
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        'End If
                    End If

                    drows = tblRegConAll.Select(str1)
                    int1 = drows.Length
                    ReDim arrRegCon(intRP + 3, int1)

                    'LEGEND: arrRegCon(intRP + 2, intRegrCt)

                    Dim dv1 As System.Data.DataView
                    dv1 = frmH.dgvAnalyticalRunSummary.DataSource
                    Dim tblARS As System.Data.DataTable = dv1.ToTable

                    Dim intCt As Short

                    intCt = 0

                    Dim intAR As Short = 0 'number of area ratio runs
                    Dim intPA As Short = 0 'number of analyte peak area runs

                    For Count3 = 1 To intRegrCt


                        var1 = tblRunID.Rows(Count3 - 1).Item("RUNID")

                        'RUNANALYTEREGRESSIONSTATUS
                        'this should not take into account rbAnalRunsShowAll
                        'it should only take into account PSAE or Accepted
                        If boolUseGroups Then
                            If frmH.chkAll.Checked Then
                                strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            ElseIf frmH.chkPSAE.Checked = False Then 'no psae
                                strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID <> 3 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            Else 'Accepted include psae
                                strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & strMatrix & "' AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            End If
                        Else
                            'If frmH.rbAnalRunsShowAll.Checked Then
                            '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            'ElseIf frmH.rbAnalRunsExclPSAE.Checked Then
                            '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            'Else 'Accepted
                            '    strF = "STUDYID = " & wStudyID & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            'End If
                        End If

                        'drows = tblRegCon.Select(strF, "REGRESSIONPARAMETERID ASC")
                        If boolExPSAE Then
                            drows = tblRegConAll.Select(strF, "REGRESSIONPARAMETERID ASC")
                        Else
                            drows = tblRegCon.Select(strF, "REGRESSIONPARAMETERID ASC")
                        End If

                        int1 = drows.Length

                        'ensure item is selected in anal run summary
                        strFARS = "[Watson Run ID] = '" & var1 & "' AND Analyte_C = '" & strAnalC & "' AND BOOLINCLUDEREGR = " & True
                        'strFARS = "[Watson Run ID] = '" & var1 & "' AND Analyte = '" & strAnal & "'"
                        Erase rowsARS
                        rowsARS = tblARS.Select(strFARS)
                        intARS = rowsARS.Length
                        'var1 = rowsARS(0).Item("BOOLINCLUDEREGR") 'DEBUG

                        If intARS = 0 Then 'skip
                        Else
                            intCt = intCt + 1
                            arrRegCon(1, intCt) = var1
                            '1=RunID, 2=RegrParameter1, etc


                            For Count4 = 0 To int1 - 1
                                var4 = NZ(drows(Count4).Item("PARAMETERVALUE"), 0)
                                var2 = var4
                                If boolLUseSigFigsRegr Then
                                    If boolLUseRegrSciNot Then
                                        var2 = Format(var4, GetScNot(LRegrSigFigs))
                                    Else
                                        'var2 = SigFigOrDec(var4, LRegrSigFigs, False)
                                        'var2 = SigFigOrDecString(var4, LRegrSigFigs, True)
                                        var2 = DisplayNum(SigFigOrDec(var4, LRegrSigFigs, False), LRegrSigFigs, False)
                                    End If
                                Else
                                    If boolLUseRegrSciNot Then
                                        'var2 = Format(var4, GetScNot(LRegrDec + 1))
                                        var2 = Format(var4, GetScNot(LRegrSigFigs))
                                    Else
                                        'var2 = Format(var4, strRegrDec)
                                        var2 = Format(var4, GetRegrDecStr(LRegrSigFigs))
                                    End If
                                End If
                                'var2 = SigFigOrDec(NZ(drows(Count4).Item("PARAMETERVALUE"), 0), LRegrSigFigs, False)
                                'var3 = Format(var2, "0.00000E-0")
                                'var3 = Format(var2, GetScNot(LRegrSigFigs))
                                var3 = var2
                                arrRegCon(Count4 + 2, intCt) = var3
                            Next

                            var4 = drows(0).Item("RSQUARED")
                            If boolLUseSigFigsRegr Then
                                If boolLUseRegrSciNot Then
                                    var2 = Format(var4, GetScNot(LR2SigFigs))
                                Else
                                    'var2 = SigFigOrDec(var4, LRegrSigFigs, False)
                                    'var2 = SigFigOrDecString(var4, LR2SigFigs, True)
                                    var2 = DisplayNum(SigFigOrDec(var4, LR2SigFigs, False), LR2SigFigs, False)
                                End If
                            Else
                                If boolLUseRegrSciNot Then
                                    var2 = Format(var4, GetScNot(LR2SigFigs))
                                Else
                                    'var2 = Format(var4, strR2Dec)
                                    var2 = Format(var4, GetRegrDecStr(LR2SigFigs))
                                End If
                            End If
                            var3 = var2
                            arrRegCon(intRP + 2, intCt) = CStr(var3) ' CStr(SigFigOrDec(var3, LR2SigFigs, False, True))
                            arrRunID(intCt) = drows(0).Item("RUNID")

                            'determine if area ratio or analyte peak areas are used
                            Dim boolAR As Boolean
                            boolAR = AreaRatioCalibr(intAnalyteID, drows(0).Item("RUNID"))
                            If boolAR Then
                                intAR = intAR + 1
                            Else
                                intPA = intPA + 1
                            End If

                        End If
                    Next

                    'determine if overall runs are area ratio or analyte peak area
                    If intAR = 0 And intPA = 0 Then
                        boolAreaRatio = True
                    Else
                        If intAR > 0 Then
                            boolAreaRatio = True
                        Else
                            boolAreaRatio = False
                        End If
                    End If

                    intTRows = intCt

                    Dim numRows As Short
                    Dim n As Short

                    numRows = 2 + (intTRows * 2) - 1
                    int1 = 0

                    If boolSTATSMEAN Then
                        int1 = int1 + 1
                    End If
                    If boolSTATSSD And boolSTATSMEAN Then
                        int1 = int1 + 1
                    End If
                    If boolSTATSCV And boolSTATSMEAN Then
                        int1 = int1 + 1
                    End If
                    If boolSTATSN Then
                        int1 = int1 + 1
                    End If
                    numRows = numRows + int1
                    numRows = numRows + 1 'for blank space
                    If BOOLINCLUDEDATE Then
                        numRows = numRows + 1 'for blank space
                    End If

                    wrdSelection = wd.Selection()


                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=numRows, NumColumns:=intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        .Selection.ParagraphFormat.RightIndent = 0

                        .Selection.Rows.AllowBreakAcrossPages = False

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

                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                        Call GlobalTableParaFormat(wd)

                        '20171220 LEE: Do not set table size, use the style default table
                        '.Selection.Font.Size = fontsize - 1
                        .Selection.Tables.Item(1).Cell(1, 1).Select()


                        .Selection.Font.Bold = False
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        '.Selection.Font.Size = 11
                        '.Selection.Tables.item(1).Columns.item(1).PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPoints
                        '.Selection.Tables.item(1).Columns.item(1).PreferredWidth = 108 'InchesToPoints(0.6)
                        ''For Count2 = 1 To intRP
                        ''    .Selection.Tables.item(1).Columns.item(Count2 + 1).PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPoints
                        ''    .Selection.Tables.item(1).Columns.item(Count2 + 1).PreferredWidth = 108 'InchesToPoints(0.6)
                        ''Next
                        '.Selection.Tables.item(1).Columns.item(intRP + 2).PreferredWidthType = Microsoft.Office.Interop.Word.wdpreferredwidthtype.wdPreferredWidthPoints
                        '.Selection.Tables.item(1).Columns.item(intRP + 2).PreferredWidth = 108 'InchesToPoints(0.6)
                        .Selection.Tables.Item(1).Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)


                        Dim strL1 As String

                        intLeg = intLeg + 1
                        strA = ChrW(intLeg + intLegStart)

                        For Count3 = 1 To intNumRegr

                            strRegressionType = arrRegrType(1, Count3)
                            'need for legend 1=Type, 2=Wt

                            'build legend
                            str1 = "Linear"
                            For Count2 = 1 To 17
                                Select Case Count2
                                    Case 1
                                        str1 = "Linear"
                                    Case 2
                                        str1 = "Isotope Dilution"
                                    Case 3
                                        str1 = "Logistic"
                                    Case 4
                                        str1 = "Quadratic"
                                    Case 5
                                        str1 = "Hyperbolic"
                                    Case 6
                                        str1 = "Burrows Watson"
                                    Case 7
                                        str1 = "Powerfit"
                                    Case 8
                                        str1 = "Logistic (Auto Estimate)"
                                    Case 9
                                        str1 = "4/5 PL"
                                    Case 10
                                        str1 = "Logit-Log"
                                    Case 11
                                        str1 = "SPLINE"
                                    Case 12
                                        str1 = "4PL"
                                    Case 13
                                        str1 = "5PL"
                                    Case 14
                                        str1 = "REGR"
                                    Case 15
                                        str1 = "Log-Log Linear"
                                    Case 16
                                        str1 = "5PL (Auto Estimate)"
                                    Case 17
                                        str1 = "Spline (Auto Smoothed)"
                                End Select
                                If StrComp(strRegressionType, str1, CompareMethod.Text) = 0 Then
                                    Exit For
                                End If
                            Next

                            'enter table heading
                            If StrComp(strRegressionType, "Quadratic", CompareMethod.Text) = 0 Then
                                str2 = "Quadratic Regression: y = Ax^2 + Bx + C"
                                str3 = "A, B, and C"
                            ElseIf StrComp(strRegressionType, "Linear", CompareMethod.Text) = 0 Then
                                str2 = "Linear Regression: y = Ax + B"
                                str3 = "A and B"
                            ElseIf StrComp(strRegressionType, "Powerfit", CompareMethod.Text) = 0 Then
                                str2 = "Powerfit Regression: Y = Ax^B"
                                str3 = "A and B"
                            Else
                                str2 = "Linear Regression: y = Ax + B"
                                str3 = "A and B"
                            End If
                            str1 = str2

                            Dim strAR As String
                            If boolAreaRatio Then
                                strAR = "peak area ratio"
                            Else
                                strAR = "analyte peak area"
                            End If

                            If boolSRegr Then
                                'str1 = str2 & " where y is the " & strAR & " of " & arrAnalytes(14, Count1) & " to Int. Std., x is the concentration of " & arrAnalytes(14, Count1) & ", and " & str3 & " are regression constants." & ChrW(10) & ChrW(9) & ChrW(9) & "Regression weighted " & strWeighting & "."
                                '20180529 LEE:
                                'Separate 'Regression' line with softreturn instead of tabs
                                str1 = str2 & " where y is the " & strAR & " of " & arrAnalytes(14, Count1) & " to Int. Std., x is the concentration of " & arrAnalytes(14, Count1) & ", and " & str3 & " are regression constants." & ChrW(11) & "Regression weighted " & strWeighting & "."
                            Else
                                str1 = str2 & " where y is the " & strAR & " of " & arrAnalytes(14, Count1) & " to Int. Std., x is the concentration of " & arrAnalytes(14, Count1) & ", and " & str3 & " are regression constants." ' Regression weighted " & strWeighting & "."
                            End If

                            '****
                            'arrBCQCs(4, Count2)
                            'search for str1 in arrLegend

                            If Count3 = 1 Then
                                strL1 = str1
                            Else
                                strL1 = strL1 & ChrW(11) & str1
                            End If

                        Next

                        arrLegend(1, intLeg) = strA
                        arrLegend(2, intLeg) = strL1
                        arrLegend(3, intLeg) = True
                        arrLegend(4, intLeg) = True
                        ctLegend = ctLegend + 1
                        'If intLeg = 1 Then '?????
                        '    arrLegend(1, intLeg) = strA
                        '    arrLegend(2, intLeg) = strL1
                        '    arrLegend(3, intLeg) = True
                        '    ctLegend = ctLegend + 1
                        'End If

                        'If boolSRegr Then
                        'Else
                        '    'add legend for RSQ
                        '    intLeg = intLeg + 1
                        '    strA = Chr(intLeg + intLegStart)
                        '    arrLegend(1, intLeg) = strA
                        '    arrLegend(2, intLeg) = "Regr. = Regression, Wt = Weighting"
                        '    arrLegend(3, intLeg) = True
                        '    arrLegend(4, intLeg) = True
                        '    ctLegend = ctLegend + 1
                        'End If


                        '***

                        If boolExcludeEntireTableTitle Then
                            intRow = 1
                        Else
                            intRow = 2
                        End If

                        strA = strAnal
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                        '***

                        'Call EnterTableNumber(wd, str2, 3)

                        'enter a table record in tblTableN
                        'ctTableN = ctTableN + 1
                        Dim dtblr As DataRow = tblTableN.NewRow
                        dtblr.BeginEdit()
                        dtblr.Item("TableNumber") = ctTableN
                        dtblr.Item("AnalyteName") = arrAnalytes(1, Count1)
                        dtblr.Item("TableName") = strTNameO
                        dtblr.Item("TableID") = intTableID
                        dtblr.Item("CHARFCID") = charFCID
                        dtblr.Item("TableNameNew") = strTName
                        tblTableN.Rows.Add(dtblr)


                        ''ensure the table is selected
                        '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToTable, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=ctTableN, Name:="")

                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)

                        '''wdd.visible = True

                        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=4, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        '.Selection.Font.Size = 11
                        .Selection.Font.Bold = False
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                        'enter headings
                        If BOOLINCLUDEDATE Then
                            '.Selection.TypeText(strWRunId & ChrW(10) & "(Analysis Date)")
                            '20180420 LEE:
                            .Selection.TypeText(strWRunId & ChrW(10) & "(" & GetAnalysisDateLabel(intTableID) & ")")
                        Else
                            .Selection.TypeText(strWRunId)
                        End If
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                        ''wdd.visible = True

                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                        For Count2 = 1 To intRP
                            str1 = ChrW(Count2 + 64) 'capital letter
                            .Selection.TypeText(str1)
                            'superscript a
                            .Selection.Font.Superscript = True
                            .Selection.TypeText(Text:=" a")
                            .Selection.Font.Superscript = False
                            '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                        Next
                        .Selection.TypeText("R-Squared")
                        '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                        If BOOLREGRULOQ Then

                            'get LLOQ and ULOQ units
                            Dim strConcUnits As String
                            Dim dv As DataView = frmH.dgvWatsonAnalRef.DataSource
                            intLeg = intLeg + 1
                            strL1 = "LLOQ: Lower Limit of Quantitation"
                            arrLegend(1, intLeg) = "b"
                            arrLegend(2, intLeg) = strL1
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True
                            ctLegend = ctLegend + 1

                            int1 = FindRowDV("LLOQ Units", dv)
                            strConcUnits = dv.Item(int1).Item(arrAnalytes(1, Count1))

                            str1 = "LLOQ"
                            .Selection.TypeText(str1)
                            .Selection.Font.Superscript = True
                            .Selection.TypeText(Text:=" b")
                            .Selection.Font.Superscript = False
                            str1 = ChrW(10) & "(" & strConcUnits & ")"
                            .Selection.TypeText(Text:=str1)
                            '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)



                            intLeg = intLeg + 1
                            strL1 = "ULOQ: Upper Limit of Quantitation"
                            arrLegend(1, intLeg) = "c"
                            arrLegend(2, intLeg) = strL1
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True
                            ctLegend = ctLegend + 1
                            int1 = FindRowDV("ULOQ Units", dv)
                            strConcUnits = dv.Item(int1).Item(arrAnalytes(1, Count1))
                            str1 = "ULOQ"
                            .Selection.TypeText(str1)
                            .Selection.Font.Superscript = True
                            .Selection.TypeText(Text:=" c")
                            .Selection.Font.Superscript = False
                            str1 = ChrW(10) & "(" & strConcUnits & ")"
                            .Selection.TypeText(Text:=str1)
                            '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                        End If

                        If boolSRegr Then
                        Else
                            var1 = "Regression"
                            .Selection.TypeText(Text:=CStr(var1))
                            '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                            ''superscript b
                            '.Selection.Font.Superscript = True
                            '.Selection.TypeText(Text:=" b")
                            '.Selection.Font.Superscript = False
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                            'enter Weighting
                            var1 = "Weighting"
                            .Selection.TypeText(Text:=CStr(var1))
                            '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                            ''superscript b
                            '.Selection.Font.Superscript = True
                            '.Selection.TypeText(Text:=" b")
                            '.Selection.Font.Superscript = False
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                        End If

                        '*****

                        'enter data in tbl
                        '1=RUNID,  2=Slope, 3=YInt, 4=R2
                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                        Count4 = 0
                        Dim intIncr As Short
                        Dim tbl As Microsoft.Office.Interop.Word.Table
                        tbl = .Selection.Tables.Item(1)
                        If boolExcludeEntireTableTitle Then
                            intIncr = 1
                        Else
                            intIncr = 2
                        End If

                        Dim vLLOQ, vULOQ

                        Dim colIncr As Short
                        If boolSTATSMEAN Or (boolSTATSSD And boolSTATSMEAN) Or (boolSTATSCV And boolSTATSMEAN) Or boolSTATSN Then
                            colIncr = 0
                        Else
                            colIncr = 0
                        End If

                        If intTRows = 0 And boolSA = False Then

                            intIncr = intIncr + 2
                            tbl.Cell(intIncr, 1).Select()
                            str1 = "There are no analytical runs that meet the criteria set in the Report Options option group shown in the StudyDoc Analytical Run Summary tab"
                            .Selection.Text = str1
                            .Selection.SelectRow()
                            .Selection.Cells.Merge()

                            GoTo end2
                        End If

                        Try
                            '20151218 LEE: Added some paragraph alignment commands to make the report look a little better

                            For Count2 = 1 To intTRows

                                strM = "Entering " & strTName & " For " & arrAnalytes(1, Count1) & " For Analytical Run # " & arrRegCon(1, Count2) & "..."
                                strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                                frmH.lblProgress.Text = strM
                                Count4 = Count4 + 1
                                intIncr = intIncr + 2
                                For Count3 = 1 To intRP + 2

                                    tbl.Cell(intIncr, Count3 + colIncr).Select()
                                    If Count3 = 1 Then 'enter run id
                                        var1 = arrRegCon(1, Count2)
                                        intRunID = var1
                                        .Selection.TypeText(CStr(var1))
                                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                                        If BOOLINCLUDEDATE Then
                                            str1 = "(" & GetDateFromRunID(NZ(var1, 0), LDateFormat, intGroup, idTR) & ")"
                                            tbl.Cell(intIncr + 1, Count3).Select()
                                            .Selection.TypeText(str1)
                                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter ' wdAlignParagraphCenter
                                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                                            tbl.Cell(intIncr, Count3).Select()
                                        End If

                                    ElseIf Count3 = intRP + 2 Then
                                        'var1 = CStr(SigFigOrDecString(arrRegCon(Count3, Count2), LRegrSigFigs, True))
                                        'do Scientific Notation instead
                                        var1 = arrRegCon(Count3, Count2)
                                        .Selection.TypeText(CStr(var1))
                                    Else
                                        var2 = NZ(arrRegCon(Count3, Count2), "NA")
                                        If IsNumeric(var2) Then
                                            'var1 = CStr(SigFigOrDec(var2, LR2SigFigs, False, True))
                                            'do Scientific Notation instead
                                            var1 = var2
                                        Else
                                            var1 = var2
                                        End If

                                        .Selection.TypeText(CStr(var1))
                                    End If

                                Next Count3

                                If BOOLREGRULOQ Then

                                    Erase rowsOQ
                                    rowsOQ = tblOQ.Select("RUNID = " & intRunID)

                                    'do LLOQ
                                    tbl.Cell(intIncr, intRP + 2 + 1).Select()
                                    If rowsOQ.Length = 0 Then
                                        vLLOQ = "NA"
                                        .Selection.TypeText(vLLOQ.ToString)
                                    Else
                                        var1 = NZ(rowsOQ(0).Item("NM"), "NA")
                                        If IsNumeric(var1) Then
                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=CStr(DisplayNum(SigFigOrDec(var1, LSigFig, False), LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Text:=CStr(Format(var1, GetRegrDecStr(LSigFig))))
                                            End If
                                        Else
                                            .Selection.TypeText(var1.ToString)
                                        End If
                                    End If

                                    'do ULOQ
                                    tbl.Cell(intIncr, intRP + 2 + 2).Select()
                                    If rowsOQ.Length = 0 Then
                                        vLLOQ = "NA"
                                        .Selection.TypeText(vLLOQ.ToString)
                                    Else
                                        var1 = NZ(rowsOQ(0).Item("VEC"), "NA")
                                        If IsNumeric(var1) Then
                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=CStr(DisplayNum(SigFigOrDec(var1, LSigFig, False), LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Text:=CStr(Format(var1, GetRegrDecStr(LSigFig))))
                                            End If
                                        Else
                                            .Selection.TypeText(var1.ToString)
                                        End If
                                    End If

                                End If

                                If boolSRegr Then
                                Else
                                    'enter Regression
                                    'tbl.Cell(intIncr, intRP + 2 + 1 + colIncr).Select()
                                    tbl.Cell(intIncr, intCols - 1).Select()
                                    var2 = arrRunID(Count2) 'returns runid
                                    var1 = GetRegrRegCon(CInt(var2), arrAnalytes(2, Count1))
                                    .Selection.TypeText(Text:=var1)

                                    'enter weighting
                                    'var1 = "1/X^2"
                                    tbl.Cell(intIncr, intRP + 2 + 2 + colIncr).Select()
                                    tbl.Cell(intIncr, intCols).Select()
                                    'var1 = strWeighting ' GetWtRegCon(CInt(var2), arrAnalytes(2, Count1))
                                    var1 = GetWtRegCon(CInt(var2), arrAnalytes(2, Count1))
                                    .Selection.TypeText(Text:=var1)
                                End If

                            Next
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try



                        n = intTRows

                        intRow = intIncr
                        If BOOLINCLUDEDATE Then
                            intRow = intRow + 1 'for blank space
                        End If

                        'enter stats
                        Dim numMean As Double
                        Dim numSD As Double
                        Dim numCV As Double
                        Dim arrA(intTRows) As Object

                        Try
                            If boolSTATSMEAN Or (boolSTATSSD And boolSTATSMEAN) Or (boolSTATSMEAN And boolSTATSCV) Or boolSTATSN Then
                                intRow = intRow + 1
                                For Count3 = 1 To intRP + 2

                                    int1 = intRow
                                    If Count3 = 1 Then
                                        If boolSTATSMEAN Then
                                            int1 = int1 + 1
                                            str1 = "Mean"
                                            tbl.Cell(int1, 1).Select()
                                            .Selection.TypeText(Text:=str1)
                                        End If
                                        If boolSTATSSD And boolSTATSMEAN Then
                                            int1 = int1 + 1
                                            str1 = "SD"
                                            tbl.Cell(int1, 1).Select()
                                            .Selection.TypeText(Text:=str1)
                                        End If
                                        If boolSTATSCV And boolSTATSMEAN Then
                                            int1 = int1 + 1
                                            str1 = ReturnPrecLabel()
                                            tbl.Cell(int1, 1).Select()
                                            .Selection.TypeText(Text:=str1)
                                        End If
                                        If boolSTATSN Then
                                            int1 = int1 + 1
                                            str1 = "n"
                                            tbl.Cell(int1, 1).Select()
                                            .Selection.TypeText(Text:=str1)
                                        End If

                                    Else

                                        int1 = intRow
                                        If boolSTATSMEAN Then
                                            int1 = int1 + 1
                                            'calculate mean
                                            var1 = 0
                                            For Count2 = 1 To intTRows
                                                var2 = CDbl(NZ(arrRegCon(Count3, Count2), 0))
                                                arrA(Count2) = var2
                                                var1 = var1 + var2
                                            Next
                                            If intTRows = 0 Then
                                                var4 = 0
                                            Else
                                                var4 = var1 / intTRows
                                            End If

                                            var2 = var4
                                            If boolLUseSigFigsRegr Then
                                                If boolLUseRegrSciNot Then
                                                    If Count3 = intRP + 2 Then 'R2
                                                        var2 = Format(var4, GetScNot(LR2SigFigs))
                                                    Else
                                                        var2 = Format(var4, GetScNot(LRegrSigFigs))
                                                    End If

                                                Else
                                                    If Count3 = intRP + 2 Then 'R2
                                                        var2 = DisplayNum(SigFigOrDec(var4, LR2SigFigs, False), LR2SigFigs, False)
                                                    Else
                                                        var2 = DisplayNum(SigFigOrDec(var4, LRegrSigFigs, False), LRegrSigFigs, False)
                                                    End If

                                                End If
                                            Else
                                                If boolLUseRegrSciNot Then
                                                    If Count3 = intRP + 2 Then 'R2
                                                        var2 = Format(var4, GetScNot(LR2SigFigs)) '+1
                                                    Else
                                                        var2 = Format(var4, GetScNot(LRegrSigFigs)) '+1
                                                    End If

                                                Else
                                                    If Count3 = intRP + 2 Then 'R2
                                                        var2 = Format(var4, GetRegrDecStr(LR2SigFigs))
                                                    Else
                                                        var2 = Format(var4, GetRegrDecStr(LRegrSigFigs))
                                                    End If
                                                End If
                                            End If
                                            numMean = CDbl(var2)
                                            str1 = CStr(var2)
                                            tbl.Cell(int1, Count3 + colIncr).Select()
                                            .Selection.TypeText(Text:=str1)

                                            If Count3 = 2 Then 'get slope
                                                Call InsertQCTables(intTableID, idTR, charFCID, -1, -1, "Slope", numMean, -1, Count1, strDo, 0, 0, False)
                                            ElseIf Count3 = intRP + 1 Then
                                                Call InsertQCTables(intTableID, idTR, charFCID, -1, -1, "YInt", numMean, -1, Count1, strDo, 0, 0, False)
                                            ElseIf Count3 = intRP + 2 Then
                                                Call InsertQCTables(intTableID, idTR, charFCID, -1, -1, "R2", numMean, -1, Count1, strDo, 0, 0, False)
                                            End If

                                        End If
                                        If boolSTATSSD And boolSTATSMEAN Then
                                            int1 = int1 + 1
                                            If intTRows < gSDMax Then
                                                var2 = "NA"
                                                numSD = 0
                                                boolNA = True
                                            Else
                                                var4 = StdDev(intTRows, arrA)
                                                var2 = var4
                                                If boolLUseSigFigsRegr Then
                                                    If boolLUseRegrSciNot Then
                                                        If Count3 = intRP + 2 Then 'R2
                                                            var2 = Format(var4, GetScNot(LR2SigFigs))
                                                        Else
                                                            var2 = Format(var4, GetScNot(LRegrSigFigs))
                                                        End If
                                                    Else
                                                        If Count3 = intRP + 2 Then 'R2
                                                            var2 = DisplayNum(SigFigOrDec(var4, LR2SigFigs, False), LR2SigFigs, False)
                                                        Else
                                                            var2 = DisplayNum(SigFigOrDec(var4, LRegrSigFigs, False), LRegrSigFigs, False)
                                                        End If
                                                    End If
                                                Else
                                                    If boolLUseRegrSciNot Then
                                                        If Count3 = intRP + 2 Then 'R2
                                                            var2 = Format(var4, GetScNot(LR2SigFigs)) '+1
                                                        Else
                                                            var2 = Format(var4, GetScNot(LRegrSigFigs)) '+1
                                                        End If

                                                    Else
                                                        If Count3 = intRP + 2 Then 'R2
                                                            var2 = Format(var4, GetRegrDecStr(LR2SigFigs))
                                                        Else
                                                            var2 = Format(var4, GetRegrDecStr(LRegrSigFigs))
                                                        End If

                                                    End If

                                                    Call InsertQCTables(intTableID, idTR, charFCID, -1, -1, "SD", numSD, -1, Count1, strDo, 0, 0, False)

                                                End If
                                                numSD = CDbl(var2)
                                            End If
                                            str1 = CStr(var2)
                                            tbl.Cell(int1, Count3 + colIncr).Select()
                                            .Selection.TypeText(Text:=str1)


                                        End If
                                        If boolSTATSCV And boolSTATSMEAN Then
                                            int1 = int1 + 1
                                            numCV = 0
                                            var2 = "NA"

                                            If intTRows < gSDMax Then
                                                numCV = 0
                                            Else
                                                If numSD = 0 Then
                                                    numCV = 0
                                                Else
                                                    numCV = numSD / numMean * 100
                                                    Call InsertQCTables(intTableID, idTR, charFCID, -1, -1, "Precision", Format(numCV, strQCDec), -1, Count1, strDo, 0, 0, False)
                                                End If
                                                var2 = numCV
                                            End If

                                            If IsNumeric(var2) Then
                                                str1 = Format(numCV, strQCDec)
                                            Else
                                                '20151218 LEE: boolNA was located right after var2="NA", resulting in an unwanted legend item
                                                'moved boolNA here where it belongs
                                                boolNA = True
                                                str1 = CStr(var2)
                                            End If

                                            tbl.Cell(int1, Count3 + colIncr).Select()
                                            .Selection.TypeText(Text:=str1)


                                        End If
                                        If boolSTATSN Then
                                            int1 = int1 + 1
                                            str1 = CStr(n)
                                            tbl.Cell(int1, Count3 + colIncr).Select()
                                            .Selection.TypeText(Text:=str1)
                                            Call InsertQCTables(intTableID, idTR, charFCID, -1, -1, "n", n, -1, Count1, strDo, 0, 0, False)
                                        End If
                                    End If
                                Next
                            End If
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try


end2:

                        Try
                            'now split table if needed
                            str1 = frmH.lblProgress.Text

                            'autofit table
                            Call AutoFitTable(wd, BOOLINCLUDEDATE)

                            'arrLegend(1, 1) = "" '"NA"
                            'arrLegend(2, 1) = strRegrLegend '"Not Applicable"
                            'arrLegend(3, 1) = False

                            If boolNA Then

                                ctLegend = ctLegend + 1
                                arrLegend(1, ctLegend) = "NA"
                                arrLegend(2, ctLegend) = "Not Applicable"
                                arrLegend(3, ctLegend) = False
                                arrLegend(4, ctLegend) = False

                            End If

                            strM = "Finalizing " & strTName & "..."
                            strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                            str1 = strM1

                            frmH.lblProgress.Text = strM1
                            frmH.Refresh()

                            str1 = strM1

                            Call SplitTable(wd, 2, ctLegend, arrLegend, str1, True, 2, False, False, False, intTableID)
                            Call MoveOneCellDown(wd)
                            Call InsertLegend(wd, intTableID, idTR, False, 1)

                            'If Count1 = frmH.arrLastAnal(2, 3) Then
                            If Count1 = ctAnalytes Then
                            Else
                                '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
                                '.Selection.TypeParagraph()
                            End If
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try


                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try

                End If

next1:

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

            Next Count1

        End With


    End Sub

    Sub MVSummaryDilutionQC_12(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

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
        Dim strMsg As String = ""

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
        Dim rowsX() As DataRow
        Dim intLegStart As Short
        Dim boolJustTable As Boolean
        Dim strTempInfo As String

        Dim int8 As Short
        Dim intExp As Short
        Dim ctExp As Short

        Dim rows2E() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False

        Dim tblD As System.Data.DataTable
        Dim intRowsDil As Short

        Dim v1, v2, vU

        Dim numPrec As Single
        Dim numTheor As Single

        Dim intGroup As Short
        Dim strAnal As String
        Dim strAnalC As String
        Dim strMatrix As String
        Dim strTNameO As String
        Dim vAnalyteIndex
        Dim vMasterAssayID
        Dim vAnalyteID
        Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted
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

        ''wdd.visible = True

        With wd

            fontsize = wd.ActiveDocument.Styles("Normal").Font.Size '.Selection.Font.Size
            fonts = fontsize '.Selection.Font.Size


            intTableID = 12

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


            strF = "IsIntStd = 'No'"
            'strS = ReturnSort(False)
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

                    intTCur = intTCur + 1

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)

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
                    strM = "Creating " & strTName & " For " & arrAnalytes(1, Count1) & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                    strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                    strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "
                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

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
                    tblLevels = dvNL.ToTable("b", True, "NOMCONC", "ALIQUOTFACTOR")
                    intNumLevels = tblLevels.Rows.Count

                    For Count2 = 0 To intNumLevels - 1 'check for any null values

                        var3 = tblLevels.Rows.Item(Count2).Item("NOMCONC")
                        If IsDBNull(var3) Then
                            str1 = "The Nominal Concentration for some assigned samples for " & var3 & " have not been configured."
                            str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                            If boolDisableWarnings Then
                            Else
                                MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            End If

                            strMsg = str1
                            'page setup according to configuration
                            str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                            'insert page break
                            'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                            Call InsertPageBreak(wd)
                            Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                            GoTo end1
                        End If
                    Next

                    'page setup according to configuration
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    ''prepare for finding dilution values
                    'tblD = dv2.ToTable("a", True, "ALIQUOTFACTOR")
                    'intRowsDil = tblD.Rows.Count

                    'find number of table rows to generate


                    Dim intRowsXTot As Short = 0

                    For Count2 = 0 To intNumRuns - 1
                        '.Selection.Tables.item(1).Cell(int1, 1).Select()
                        'enter runid
                        var10 = tblNumRuns.Rows(Count2).Item("RUNID")

                        '.Selection.TypeText(CStr(var10))
                        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                        intRowsX = 0
                        For Count3 = 0 To intNumLevels - 1
                            varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                            var1 = tblLevels.Rows(Count3).Item("ALIQUOTFACTOR")
                            dv2.RowFilter = ""
                            'don't know why, but must make a long filter here or
                            'both analytes get returned in dv2.rowfilter
                            strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ALIQUOTFACTOR = " & var1
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
                    intTblRows = intTblRows + (1 * intNumRuns) 'for a blank row after each run set
                    'intTblRows = intTblRows + (5 * intNumRuns) 'for Mean/Bias/n section for each run set

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + intCSN

                    If intCSN > 0 Then
                        intTblRows = intTblRows + (1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                    Else
                        intTblRows = intTblRows - 1 'subtract an unneeded blank row
                    End If

                    If boolQCREPORTACCVALUES Then

                    Else
                        'intTblRows = intTblRows + 3 'for stats headings
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

                    '20180221 LEE:
                    'Alturas reports that sometimes Diln samples are cut off at the botton.
                    'This means not enough rows are being produced
                    'Add 20 rows at end to make sure
                    intTblRows = intTblRows + 20

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

                            'wd.Visible = True

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
                        For Count2 = 0 To intNumLevels - 1
                            'var1 = arrBCQCs(3, Count2)
                            If boolLUseSigFigs Then
                                'var1 = CStr(SigFigOrDecString(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig, False))
                                var1 = DisplayNum(SigFigOrDec(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig, False), LSigFig, False)
                            Else
                                var1 = Format(RoundToDecimalRAFZ(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig))
                            End If


                            If LboolNomConcParen Then
                                var1 = "QC Dilution" & ChrW(10) & "(" & var1 & ChrW(160) & strConcUnits & ")"
                            Else
                                var1 = "QC Dilution " & ChrW(10) & var1 & ChrW(160) & strConcUnits
                            End If

                            .Selection.TypeText(Text:=var1)
                            intLeg = intLeg + 1
                            strA = Chr(intLeg + 96)

                            'enter superscript
                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                            'record two legend items
                            arrLegend(1, intLeg) = strA
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True
                            ctLegend = ctLegend + 1
                            'record legend entry for arrLegend(2,n) later in code

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

                        Dim intDL As Short = 0
                        int1 = 4 'row position counter
                        For Count2 = 0 To intNumRuns - 1  'For all runs represented in the Assigned Samples list...

                            .Selection.Tables.Item(1).Cell(int1, 1).Select()
                            'enter runid
                            var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")

                            'strM = "Creating Summary of Interpolated Dilution QC Standard Concentrations Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
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
                            boolOutHeadE = False
                            boolOutHeadI = False
                            boolDeleteRows = False
                            Dim varDil
                            Dim int12 As Short
                            Dim boolEnterDiff As Boolean
                            int12 = -1
                            For Count3 = 0 To (intNumLevels * int11) - 1 Step int11  'For all nominal concentrations, for all aliquot factors

                                boolEnterDiff = False
                                int12 = int12 + 1
                                varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                                varDil = tblLevels.Rows.Item(Count3).Item("ALIQUOTFACTOR")
                                'determine hi and lo (nom*flagpercent)
                                strF = "CONCENTRATION = '" & varNom & "'"
                                'rows10 = tblBCQCs.Select(strF)

                                'determine hi and lo (nom*flagpercent)
                                'strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10

                                'don't need masterassayid and analyteindex anymore
                                strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10

                                'if Conc < 1, then the query return 0 records
                                'must do something different
                                var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)

                                'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                                arrFP(1, Count3) = var1
                                arrFP(2, Count3) = var1
                                v1 = var1
                                v2 = var1
                                vU = 0

                                Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)


                                '20160816 LEE: intRowsX needs to be determined for each dataset
                                '****
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter


                                'var10
                                intRowsX = 0
                                int12 = -1
                                For Count4 = 0 To tblLevels.Rows.Count - 1
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

                                'start entering data
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                'strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom
                                strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ALIQUOTFACTOR = " & varDil
                                dv2.RowFilter = strF
                                int2 = dv2.Count


                                If int2 > 0 Then      'Some Runs will *not* have all Nominal Concentrations or all Aliquot Factors (dv2.Count = int2 = 0) 
                                    'create rows1 from tbl1 which will contain data
                                    strF = ""
                                    'Erase rows2
                                    'rows2 = tbl1.Select(strF)
                                    Dim tbl2s As System.Data.DataTable = dv2.ToTable
                                    rows2 = tbl2s.Select("RUNID > -1")
                                    int3 = rows2.Length

                                    'int3 = rows2.Length
                                    nI = rows2.Length

                                    'redo hi/lo
                                    vU = rows2(0).Item("BOOLUSEGUWUACCCRIT")
                                    If vU = -1 Then
                                        v1 = CDec(NZ(rows2(0).Item("NUMMAXACCCRIT"), 0))
                                        v2 = CDec(NZ(rows2(0).Item("NUMMINACCCRIT"), 0))
                                        arrFP(1, Count3) = v1
                                        arrFP(2, Count3) = v2

                                        Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)

                                    End If

                                    '''''''''''''console.writeline(strF)

                                    'now do excluded
                                    strF = ""

                                    Erase rows2E
                                    'rows2E = tbl1.Select(strF)
                                    If gAllowExclSamples And LAllowExclSamples Then
                                        rows2E = tbl2s.Select("ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0")
                                    Else
                                        rows2E = tbl2s.Select("ELIMINATEDFLAG = 'N'")
                                    End If
                                    nE = rows2E.Length

                                    'record legend entry
                                    'get dilution factor
                                    var1 = rows2(int12).Item("ALIQUOTFACTOR")
                                    var1 = 1 / var1
                                    var3 = VerboseNumber(var1, False)
                                    str1 = "Dilution QCs undiluted concentration " & varNom & " " & strConcUnits
                                    numDF = rows2(0).Item("ALIQUOTFACTOR")
                                    var3 = 1 / numDF
                                    'var3 = RoundToDecimal(var3, 0)
                                    var3 = GetDilnFactor(CDec(var3)) '20190220 LEE
                                    Dim strAN As String = GetAN(var3)

                                    str1 = str1 & "; " & strAN & " " & var3 & "-fold dilution with blank plasma was done prior to extraction and analysis."
                                    intDL = intDL + 1
                                    arrLegend(2, intDL) = str1

                                    'enter data
                                    For Count4 = 0 To intRowsX - 1 'int3 - 1

                                        boolOC = False

                                        .Selection.Tables.Item(1).Cell(int1 + Count4, Count3 + 2).Select()
                                        If Count4 > nI - 1 Then
                                            'str1 = "NA"

                                            If boolQCNA Then
                                                str1 = "NA"
                                            Else
                                                str1 = ""
                                            End If

                                            .Selection.TypeText(str1)
                                            boolEnterDiff = False
                                            boolOC = True

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

                                                'Remember, tblAssignedSamples does not have DECISIONREASON
                                                Dim var6
                                                var6 = "No Value: " & GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))

                                                boolEnterDiff = True 'FALSE
                                                intExp = intExp + 1
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                '20160305 LEE:
                                                'Added DECISIONREASON code
                                                'Set Legend String
                                                str1 = GetLegendStringExcluded(arrFP(1, int12), arrFP(2, int12), vU, var6, intTableID, True, "")
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                fonts = .Selection.Font.Size
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

                                                'Remember, tblAssignedSamples does not have DECISIONREASON
                                                Dim var6
                                                var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))

                                                boolEnterDiff = True 'FALSE
                                                intExp = intExp + 1
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                '20160305 LEE:
                                                'Added DECISIONREASON code
                                                'Set Legend String
                                                str1 = GetLegendStringExcluded(arrFP(1, int12), arrFP(2, int12), vU, var6, intTableID, True, "")
                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                fonts = .Selection.Font.Size
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=DisplayNum(var2, LSigFig, False))
                                                Else
                                                    .Selection.TypeText(Text:=Format(var2, GetRegrDecStr(LSigFig)))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                            Else

                                                boolEnterDiff = True
                                                'determine if value is outside acceptance criteria
                                                'If var2 > hi Or var2 < lo Then 'flag
                                                If OutsideAccCrit(var2, varNom, v1, v2, NZ(vU, 0)) Then
                                                    intLeg = intLeg + 1
                                                    strA = ChrW(intLeg + intLegStart)

                                                    'Set Legend String
                                                    str1 = GetLegendStringIncluded(arrFP(1, int12), arrFP(2, int12), vU)
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
                                                Else
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=CStr(Format(var2, GetRegrDecStr(LSigFig))))
                                                    End If
                                                End If
                                            End If

                                        End If

                                        If boolSTATSDIFFCOL Then
                                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                            If boolEnterDiff Then
                                                'var3 = Format(((var2 / varNom) - 1) * 100, strQCDec)
                                                If boolTHEORETICAL Then
                                                    var3 = CalcREPercent(var2, varNom, intQCDec)
                                                    numTheor = 100 + CDec(var3)

                                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numTheor, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                                Else
                                                    var3 = Format(RoundToDecimal(CalcREPercent(var2, varNom, intQCDec), intQCDec), strQCDec)

                                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", var3, CSng(var10), Count1, strDo, v1, v2, boolOC)

                                                End If
                                            Else
                                                'var3 = "NA"

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
                                End If

                                'now enter Mean/Bias/n
                                If Count3 = 0 Then

                                    int8 = 0

                                    '''''''''''wdd.visible = True

                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        If intExp = 0 Then
                                        Else
                                            int8 = int8 + 1
                                            If boolOutHeadE Then
                                            Else
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 2).Select()
                                                .Selection.TypeText(Text:="Summary Statistics Excluding Outlier Values")

                                                Try
                                                    .Selection.Cells.Merge()
                                                Catch ex As Exception

                                                End Try
                                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                                Try
                                                Catch ex As Exception

                                                End Try

                                                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                                End With

                                            End If


                                        End If
                                    End If

                                    If Count3 = 0 Then
                                        Call typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)

                                    End If

                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        If intExp = 0 Then
                                        Else

                                            int8 = int8 + 2

                                            If boolOutHeadI Then
                                            Else
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 2).Select()
                                                .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                                Try
                                                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                                    .Selection.Cells.Merge()
                                                Catch ex As Exception

                                                End Try

                                                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                                End With

                                                If Count3 = 0 Then
                                                    Call typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                                int8 = 0
                                If int2 > 0 Then
                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        If intExp = 0 Then
                                        Else
                                            int8 = int8 + 1
                                        End If
                                    End If

                                    v1 = arrFP(1, int12)
                                    v2 = arrFP(2, int12)


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
                                                        .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                    End If

                                                    Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                    boolEnterDiff = True
                                                Else
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                    End If
                                                    boolEnterDiff = True
                                                End If
                                            End If

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

                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numSD, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(numSD, GetRegrDecStr(LSigFig))))
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
                                    End If
                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        Try
                                            'enter %Bias
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                            If nE = 0 Then
                                                .Selection.TypeText("NA")
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
                                    End If

                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        Try
                                            'enter %theoretical
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                            If nE = 0 Then
                                                .Selection.TypeText("NA")
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
                                    End If

                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        Try
                                            'enter %Diff
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                            If nE = 0 Then
                                                .Selection.TypeText("NA")
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
                                    End If
                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        Try
                                            'enter %RE
                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                            If nE = 0 Then
                                                .Selection.TypeText("NA")
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
                                        If intExp = 0 Then
                                        Else
                                            int8 = int8 + 2

                                            If boolSTATSMEAN Then
                                                Try
                                                    'enter Mean
                                                    int8 = int8 + 1
                                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                    var1 = MeanDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                                    If boolLUseSigFigs Then
                                                        numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                    Else
                                                        numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If

                                                    '.Selection.TypeText(CStr(numMean))

                                                    If nI = 0 Then
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
                                                                .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                            End If

                                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                            boolEnterDiff = True
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                            Else
                                                                .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                            End If

                                                            boolEnterDiff = True
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
                                                            .Selection.TypeText(Text:=CStr(DisplayNum(numSD, LSigFig, False)))
                                                        Else
                                                            .Selection.TypeText(Text:=CStr(Format(numSD, GetRegrDecStr(LSigFig))))
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
                                                    If nI = 0 Then
                                                        .Selection.TypeText("NA")
                                                    Else
                                                        .Selection.TypeText(Format(numBias, strQCDec))
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

                                                    numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                                    numTheor = 100 + CDec(numTheor)
                                                    If nI = 0 Then
                                                        .Selection.TypeText("NA")
                                                    Else
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
                                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                                    If nI = 0 Then
                                                        .Selection.TypeText("NA")
                                                    Else
                                                        .Selection.TypeText(Format(numBias, strQCDec))
                                                    End If

                                                Catch ex As Exception

                                                End Try
                                            End If

                                            If BOOLSTATSRE And boolSTATSMEAN Then
                                                Try
                                                    'enter %diff
                                                    int8 = int8 + 1
                                                    .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                    numBias = CalcREPercent(numMean, varNom, intQCDec)
                                                    If nI = 0 Then
                                                        .Selection.TypeText("NA")
                                                    Else
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
                                Else
                                    'NDL: Still need to update int8 to keep everything in the right place.  There is most likely a cleaner way to do this.
                                    int8 = int8 - boolQCREPORTACCVALUES - boolSTATSMEAN - boolSTATSSD - boolSTATSCV - boolSTATSBIAS - _
                                        boolTHEORETICAL - boolSTATSDIFF - BOOLSTATSRE
                                    If (boolQCREPORTACCVALUES) Then
                                    Else
                                        int8 = int8 + 2 - boolSTATSMEAN - boolSTATSSD - boolSTATSCV - boolTHEORETICAL - boolSTATSDIFF - BOOLSTATSRE - boolSTATSN
                                    End If
                                End If
                            Next

                            'increase row position counter
                            If Count2 >= (intNumRuns * int11) - 1 Then
                                int1 = int1 + intRowsX + int8 + 1
                            Else
                                int1 = int1 + intRowsX + int8 + 2
                            End If
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
                        '.Selection.Cells.Merge()
                        If intNumLevels < 2 And boolSTATSDIFFCOL = False Then
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


                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    'go to end of table
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn)

                    'enter table number
                    str1 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated Dilution QC Standard Concentrations"

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
                    Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                    '***

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

                    str1 = strM1

                    Call SplitTable(wd, 3, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'move to line below table
                    '.Selection.Tables.item(1).Select()
                    '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)

                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)

                    'wd.Visible = True
                    'MsgBox("DilQC")
                    'wd.Visible = False


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
                        var1 = var1 'debug
                    End If

                End If

next1:


            Next
end2:
        End With

    End Sub

    Sub MVSummaryTempStabilityQC_18(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)

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
        '1=max, 2=min
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

        Dim intGroup As Short
        Dim strAnal As String
        Dim strAnalC As String
        Dim strMatrix As String
        Dim strTNameO As String
        Dim intRunID As Int16
        Dim strDECISIONREASON As String
        Dim boolExFromAS As Boolean

        Dim v1, v2, vU

        Dim numPrec As Single
        Dim numTheor As Single

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size 'wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        With wd

            intTableID = 18

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
            strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies
            rowsX = tbl2.Select(strF)


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
                        strM = "Creating Summary of " & strTempInfo & " Temperature Stability in Matrix Table ...."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strM = "Creating " & strTName & "..."
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        strM1 = strM
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
                    strM = "Creating Summary of " & strTempInfo & " Temperature Stability in Matrix Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
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

                    'determine if any outliers
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
                    Erase rows2
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug

                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count

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
                    intTblRows = intTblRows + 2 'for header
                    intTblRows = intTblRows + 1 'for blank row
                    intTblRows = intTblRows + (intRowsXTot) 'for number of data rows
                    intTblRows = intTblRows + (1 * intNumRuns) 'for a blank row after each run set
                    'intTblRows = intTblRows + (5 * intNumRuns) 'for Mean/Bias/n section for each run set

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    intTblRows = intTblRows + (intCSN * intNumRuns)

                    If intCSN > 0 Then
                        intTblRows = intTblRows + (1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set
                    Else
                        'intTblRows = intTblRows - 1 'subtract an unneeded blank row
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


                    Try

                        '20180511 LEE:
                        'investigate here
                        'wd.Visible = True

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
                        For Count2 = 0 To intNumLevels - 1

                            'var1 = arrBCQCs(3, Count2)
                            If boolLUseSigFigs Then
                                var1 = DisplayNum(SigFigOrDec(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig, False), LSigFig, False)
                            Else
                                var1 = Format(RoundToDecimalRAFZ(tblLevels.Rows.Item(Count2).Item("NOMCONC"), LSigFig), GetRegrDecStr(LSigFig))
                            End If

                            var2 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                            '.Selection.TypeText(Text:=var3)

                            '******determine if the level is a diln level
                            Dim strE As String
                            'var3 = var2 ' & ChrW(10) & var1 & " " & strConcUnits
                            var3 = ReturnStdQC(var2.ToString)
                            If LboolNomConcParen Then
                                strE = ChrW(10) & "(" & var1 & ChrW(160) & strConcUnits & ")"
                            Else
                                strE = ChrW(10) & var1 & ChrW(160) & strConcUnits
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
                                    'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & SigFigOrDecString(CDbl(var1), LSigFig, True) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                    arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & SigFigOrDecString(CDbl(var1), LSigFig, True) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
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

                            'strM = "Creating Summary of " & strTempInfo & " Stability in Matrix Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
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

                            Dim int12 As Short
                            Dim boolEnterDiff As Boolean


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
                                v1 = var1
                                v2 = var1
                                vU = 0

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

                                'create rows1 from tbl1 which will contain data
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
                                nI = int3

                                If int3 = 0 Then
                                Else

                                    'redo hi/lo
                                    vU = rows2(0).Item("BOOLUSEGUWUACCCRIT")
                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                        v1 = CDec(NZ(rows2(0).Item("NUMMAXACCCRIT"), 0))
                                        v2 = CDec(NZ(rows2(0).Item("NUMMINACCCRIT"), 0))
                                        arrFP(1, int12) = v1
                                        arrFP(2, int12) = v2

                                        Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)

                                    End If

                                End If

                                'get included dataset
                                strF = ""
                                'For Count4 = 0 To dv2.Count - 1
                                '    var2 = dv2(Count4).Item("ANALYTEINDEX")
                                '    var3 = dv2(Count4).Item("MASTERASSAYID")
                                '    var4 = dv2(Count4).Item("RUNSAMPLEORDERNUMBER")
                                '    var5 = dv2(Count4).Item("ANALYTEID")

                                '    If Count4 <> dv2.Count - 1 Then
                                '        strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var5 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLEORDERNUMBER = " & var4 & " AND ELIMINATEDFLAG = 'N') OR "
                                '    Else
                                '        strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var5 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLEORDERNUMBER = " & var4 & " AND ELIMINATEDFLAG = 'N')"
                                '    End If
                                'Next
                                Erase rows2E
                                'rows2E = tbl1.Select(strF)
                                If gAllowExclSamples And LAllowExclSamples Then
                                    rows2E = tbl2s.Select("ELIMINATEDFLAG = 'N' AND BOOLEXCLSAMPLE = 0")
                                Else
                                    rows2E = tbl2s.Select("ELIMINATEDFLAG = 'N'")
                                End If

                                nE = rows2E.Length

                                'enter data
                                boolEnterDiff = False
                                For Count4 = 0 To intRowsX - 1 'int3 - 1

                                    boolOC = False

                                    .Selection.Tables.Item(1).Cell(int1 + Count4, Count3 + 2).Select()
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

                                            boolEnterDiff = True 'FALSE
                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            strA = ChrW(intLeg + intLegStart)

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            Dim var6
                                            'Remember, tblAssignedSamples does not have DECISIONREASON
                                            var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))
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
                                        Else
                                            boolEnterDiff = True
                                            'determine if value is outside acceptance criteria
                                            'If var2 > hi Or var2 < lo Then 'flag
                                            If OutsideAccCrit(var2, varNom, v1, v2, NZ(vU, 0)) Then
                                                intLeg = intLeg + 1
                                                strA = ChrW(intLeg + intLegStart)

                                                'Set Legend String
                                                str1 = GetLegendStringIncluded(arrFP(1, int12), arrFP(2, int12), vU)
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
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=DisplayNum(var2, LSigFig, False))
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


                                'now enter Mean/Bias/n
                                int8 = 0

                                If boolQCREPORTACCVALUES Then
                                Else
                                    If boolOutlier Then
                                        int8 = int8 + 1
                                    End If
                                End If


                                If Count3 = 0 Then
                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        If boolOutlier Then
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 2).Select()
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

                                    'Type Text labels for Statistics
                                    Call typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)

                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        If boolOutlier Then
                                            int8 = int8 + 2
                                            .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, 2).Select()
                                            .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                            Try
                                                .Selection.Cells.Merge()
                                            Catch ex As Exception

                                            End Try
                                            With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                            End With

                                            'Type Text labels for Statistics with Outliers
                                            Call typeStatsLabels(wd, int8, int1 + intRowsX, 1, False)
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

                                intStart = int8

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
                                            boolMean = False
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
                                                    .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                boolEnterDiff = True
                                            Else
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If
                                                boolEnterDiff = True
                                            End If
                                        End If


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

                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=CStr(DisplayNum(numSD, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Text:=CStr(Format(numSD, GetRegrDecStr(LSigFig))))
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
                                        End If
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolSTATSBIAS And boolSTATSMEAN Then
                                    Try
                                        'enter %Bias
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                        If boolMean Then
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        Else
                                            .Selection.TypeText("NA")
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
                                End If

                                If boolTHEORETICAL And boolSTATSMEAN Then
                                    Try
                                        'enter %theoretical
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                        If boolMean Then
                                            .Selection.TypeText(Format(numTheor, strQCDec))
                                        Else
                                            .Selection.TypeText("NA")
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
                                End If
                                If boolSTATSDIFF And boolSTATSMEAN Then
                                    Try
                                        'enter %Bias
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                        If boolMean Then
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        Else
                                            .Selection.TypeText("NA")
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
                                End If
                                If BOOLSTATSRE And boolSTATSMEAN Then
                                    Try
                                        'enter %RE
                                        int8 = int8 + 1
                                        .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()

                                        If boolMean Then
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        Else
                                            .Selection.TypeText("NA")
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
                                                var1 = MeanDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                                If boolLUseSigFigs Then
                                                    numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                Else
                                                    numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                                End If

                                                '.Selection.TypeText(CStr(numMean))

                                                'determine if value is outside acceptance criteria
                                                'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                                If nI = 0 Then
                                                    .Selection.TypeText("NA")
                                                    boolMean = False
                                                Else

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
                                                            .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                        End If

                                                        Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                        boolEnterDiff = True
                                                    Else
                                                        If boolLUseSigFigs Then
                                                            .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                        Else
                                                            .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                        End If
                                                        boolEnterDiff = True
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
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(numSD, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=CStr(Format(numSD, GetRegrDecStr(LSigFig))))
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

                                                numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                                numTheor = 100 + CDec(numTheor)

                                                If boolMean Then
                                                    .Selection.TypeText(Format(numTheor, strQCDec))
                                                Else
                                                    .Selection.TypeText("NA")
                                                End If

                                            Catch ex As Exception

                                            End Try

                                        End If

                                        If boolSTATSDIFF And boolSTATSMEAN Then
                                            Try
                                                'enter %Bias
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                numBias = CalcREPercent(numMean, varNom, intQCDec)

                                                If boolMean Then
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                Else
                                                    .Selection.TypeText("NA")
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If BOOLSTATSRE And boolSTATSMEAN Then
                                            Try
                                                'enter %RE
                                                int8 = int8 + 1
                                                .Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                                numBias = CalcREPercent(numMean, varNom, intQCDec)

                                                If boolMean Then
                                                    .Selection.TypeText(Format(numBias, strQCDec))
                                                Else
                                                    .Selection.TypeText("NA")
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
                            Next

                            'increase row position counter
                            If Count2 = intNumRuns - 1 Then
                                int1 = int1 + intRowsX + int8 + 1
                            Else
                                int1 = int1 + intRowsX + int8 + 2
                            End If
                        Next

                        'bottom border this row
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        Call RemoveRows(wd, 1)

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
                        ' Try
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


                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    'go to end of table
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn)

                    'enter table number
                    var1 = strTempInfo
                    str2 = var1
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

                    'str1 = str2 & " Stock Solution Stability Assessment: Summary of " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")
                    'str1 = "Summary of " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " " & strTempInfo & " Stability in Matrix"
                    str1 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " " & str2 & " Stability in Matrix"

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

                    'autofit table
                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    str1 = strM1

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'autofit table
                    Call AutoFitTable(wd, False)

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
                '    'str2 = "Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " " & str2 & " Stability in Matrix"
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
        End With

    End Sub


    Sub MVAdHocQCStabilityComparison_32(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64, intTID As Short)


        '20190118 LEE:
        'Note that the following tables are now redirected here:
        '22	[Period Temp] Stock Solution Stability Assessment
        '23	[Period Temp] Spiking Solution Stability Assessment
        '29	[Period Temp] Long-Term QC Std Storage Stability

        'if redirected is 29 Long-Term, boolRCConc has been forced to TRUE at modReport.PrepareTable


        'based on final extract stability

        'NOTES:

        'In the Ad Hoc Stability Comparison table, the data sets are sorted and listed according to the following ordering ‘sequence:
        '
        '	Run ID (ascending by numeric)
        '	Caption (ascending by alphabetical)
        '
        'The comparison value (Difference, Recovery, Mean Accuracy) is calculated using this logic:
        '
        'Positive Results  = 
        '	[First Data Set]*100/[Second Data Set]  =  
        '	[New Data]*100/[Old Data] = 
        '	[QCNL]*100/[QCSL]
        '
        'Negative Results  =  
        '	[Second Data Set]*100/[First Data Set]  =  
        '	[Old Data]*100/[New Data] = 
        '	[QCSL]*100/[QCNL]

        '20180309 LEE:
        'Added table logic to put data sets in columns, rather than rows
        'This will allow users to enter several time points
        'First data set configured must be T0
        '%Diff will be row after stats
        'For normal display, will show %Diff in a column same row as mean rather than following stats

        '20181220 LEE:
        'Added 'No Calculations' checkbox
        'Some clients want to show analytical run data with stats for each run, but don't want to do any comparison
        'The global variable will be BOOLCSREPORTACCVALUES, which is deprecated from Cal Std options


        Dim boolOC As Boolean = False 'bool if eliminated
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
        Dim Count6 As Short
        Dim Count10 As Short
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
        Dim boolPro As Boolean

        Dim boolHasOutlier As Boolean = False

        Dim hi, lo
        Dim rows10() As DataRow
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim arrFP(2, 20) 'FlagPercent array
        '1=max, 2=min
        Dim strFP As String
        Dim numMean As Decimal
        Dim numMean1 As Decimal
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
        Dim rows2A() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False
        Dim arrS(10, 1)
        Dim boolFirst As Boolean = False
        Dim numInj As Short

        Dim tblRID As System.Data.DataTable
        Dim numRID As Short
        Dim ctTbl As Short
        Dim varAnal, varIS

        Dim strCombTitle As String = ""
        Dim strCombTitle1 As String = ""
        Dim strCombTitle2 As String = ""

        Dim varConc

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

        Dim numT01 As Single
        Dim numT02 As Single 'for excluded sample calculation

        Dim int12 As Short = -1
        Dim boolNA As Boolean

        Dim intDiffRow1 As Short
        Dim intDiffRow2 As Short

        '20181127 LEE:
        Dim intBSNR As Short = 1 'boolStatsNR. Used for Stock (8) and Spiking (9) solution
        intBSNR = GetStatsNR(idTR)
        Dim boolBSNR As Boolean
        If intBSNR = 8 Or intBSNR = 9 Then
            boolBSNR = True
        Else
            boolBSNR = False
        End If

        '20190225 LEE:
        Dim strDiff As String = "%Difference"

        'boolAdHocStabCompColumns:  This is the grid-style to show sets of data
        Dim boolColDiff As Boolean = Not (boolAdHocStabCompColumns) 'if data is NOT to be shown in columns, then report %Diff in it's own column, not after stats section
        '20181219 LEE:
        'Keep boolColDiff logic in place, look at boolNoneLeg

        '20180724 LEE:
        'New Logic: if no legend, then no diff columns
        'If boolNONELEG Then
        '    boolColDiff = False
        'End If
        '20181220 LEE:
        'new boolean for boolColDiff
        If boolCSREPORTACCVALUES Then
            boolColDiff = False
        Else
            boolColDiff = True
        End If

        Dim charFCID As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        boolJustTable = False

        Cursor.Current = Cursors.WaitCursor

        ''wdd.visible = True

        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size 'wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        'NOTE: This table expects separation of data based on separate Watson Run ID's
        'it should also separate based on Run Identifier

        With wd

            intTableID = intTID '32

            Dim strWRunId As String = GetWatsonColH(intTableID)
            Dim strLabel As String = GetLabelColH(intTableID)

            Dim boolSampleNameCol As Boolean = False
            Dim strSampleNameCol As String = GetSampleName(intTableID)

            If boolAdHocStabCompColumns Then
                boolSampleNameCol = False
            Else
                If Len(strSampleNameCol) = 0 Then
                    boolSampleNameCol = False
                Else
                    boolSampleNameCol = True
                End If
            End If


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

            'this table may include intstd

            strF = "IsIntStd = 'No'"
            If boolIncludeISTbl Then
                strF = "IsIntStd = 'No' OR IsIntStd = 'Yes'"
            End If
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
                boolNA = False

                '20180807 LEE:
                'Create table to store means for %Diff reporting
                Dim tblMeans As New DataTable
                Dim col1 As New DataColumn
                col1.ColumnName = "boolAll"
                col1.DataType = System.Type.GetType("System.Int16")
                tblMeans.Columns.Add(col1)

                Dim col2 As New DataColumn
                col2.ColumnName = "Mean"
                col2.DataType = System.Type.GetType("System.Single")
                tblMeans.Columns.Add(col2)

                Dim col3 As New DataColumn
                col3.ColumnName = "Run"
                col3.DataType = System.Type.GetType("System.Int16")
                tblMeans.Columns.Add(col3)

                Dim col4 As New DataColumn
                col4.ColumnName = "Level"
                col4.DataType = System.Type.GetType("System.Int16")
                tblMeans.Columns.Add(col4)

                Dim rowsMeanAll() As DataRow 'T0 all
                Dim rowsMeanAll1() As DataRow
                Dim rowsMean() As DataRow
                Dim rowsMean1() As DataRow
                Dim numMeanMT0 As Decimal
                Dim numMeanMAllT0 As Decimal
                Dim numMeanMAll As Decimal
                Dim numMeanM As Decimal
                Dim boolHasOut As Boolean = False
                Dim intRunM As Short = -1

                '20190117 LEE
                Dim rowsMeanT0All() As DataRow 'T0 accepted
                Dim numMeanT0All As Decimal

                'check if table is to be generated
                'strDo = arrAnalytes(1, Count1) 'record column name
                Dim strX As String
                strDo = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")

                'If boolIncludeISTbl Then
                'Else
                '    If UseAnalyte(CStr(strDo)) Then
                '    Else
                '        GoTo next1
                '    End If
                'End If



                strX = rows11(Count1 - 1).Item("IsIntStd")
                'bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                If boolIncludeISTbl And StrComp(strX, "Yes", CompareMethod.Text) = 0 Then
                    bool = True
                Else
                    bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                End If


                Dim strM1 As String
                Dim intLRow As Short
                If bool Then 'continue

                    ctTbl = ctTbl + 1

                    intTCur = intTCur + 1

                    Dim intCRow As Short = 2

                    If boolBSNR Then
                        intCRow = 1
                    End If

                    intLRow = intCRow

                    '****

                    'check if table is to be generated
                    'strDo = arrAnalytes(1, Count1) 'record column name
                    Dim boolX As Boolean

                    strDo = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")
                    strX = rows11(Count1 - 1).Item("IsIntStd")
                    boolX = False

                    'var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    'var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    'var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    'vAnalyteIndex = var1
                    'vMasterAssayID = var2
                    'vAnalyteID = tbl4.Rows.Item(Count1 - 1).Item("AnalyteID")

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

                    '****

                    'ensure data has been entered
                    strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strDo) & "' AND ID_TBLREPORTTABLE = " & idTR
                    rowsX = tbl2.Select(strF)
                    If rowsX.Length = 0 Then



                        strM = "Creating Summary of " & strTempInfo & " Final Extract Stability Table ...."
                        strM = "Creating " & strTName & "..."
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
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    Dim strISConc As String = ""

                    If StrComp(strX, "Yes", CompareMethod.Text) = 0 Then
                        If boolIncludeISTbl Then
                            'get strISConc
                            Dim strFIS As String
                            Dim rowsFIS() As DataRow
                            Dim tblP As System.Data.DataTable = tblTableProperties
                            strFIS = "ID_TBLREPORTTABLE = " & idTR
                            rowsFIS = tblP.Select(strFIS)
                            If boolLUseSigFigs Then
                                strISConc = NZ(rowsFIS(0).Item("CHARISCONC"), "")
                            Else
                                strISConc = NZ(rowsFIS(0).Item("CHARISCONC"), "")
                            End If
                        Else
                            GoTo next1
                        End If
                    Else
                        If UseAnalyte(CStr(strDo)) Then
                        Else
                            GoTo next1
                        End If
                    End If


                    strM = "Creating " & strTName & " For " & arrAnalytes(1, Count1) & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    '*****

                    If StrComp(strX, "Yes", CompareMethod.Text) = 0 Then
                        'check for boolIntStd in tbl2
                        strF = "IsIntStd = 'Yes'"
                        strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "BOOLINTSTD = -1"
                        Erase rowsX
                        rowsX = tbl2.Select(strF2)
                        int1 = rowsX.Length
                        boolX = True

                        If int1 > 0 And boolIncludeISTbl Then
                            bool = True

                        Else
                            bool = True 'False
                        End If
                    Else
                        bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                    End If


                    '****
                    If boolX Then
                        var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        strF2 = "BOOLINTSTD = -1 AND "
                        strF2 = strF2 & "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "CHARANALYTE = '" & CleanText(cstr(var3)) & "'"
                    Else
                        strF2 = "BOOLINTSTD = 0 AND ID_TBLSTUDIES = " & id_tblStudies & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "INTGROUP = " & intGroup

                    End If
                    '****

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    strS = "RUNID ASC, CHARHELPER2 ASC, RUNSAMPLEORDERNUMBER ASC"
                    '20180309 LEE:
                    'remove CHARHELPER2 from sort
                    'need to keep CHARHELPER2 in SampleOrderNumber
                    'users must ensure T0 samples are acquired first
                    'must be order in which samples were assigned
                    strS = "ID_TBLASSIGNEDSAMPLES ASC" ' "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug


                    'find number of charhelper2's
                    Dim tblH2 As System.Data.DataTable = dv2.ToTable("H2", True, "CHARHELPER2")

                    Dim numH2 As Short
                    numH2 = tblH2.Rows.Count

                    Dim numH1 As Short
                    Dim tblH1 As System.Data.DataTable

                    '20181127 LEE: Account for Stock (8) or Spiking (9) solution
                    If boolBSNR Then
                        str1 = "NOMCONC"
                    Else
                        'find number of charhelper1's - this would be levels
                        'Dim tblH1 As System.Data.DataTable = dv2.ToTable("H1", True, "CHARHELPER1")
                        '20180823 LEE:
                        If INTQCLEVELGROUP = 0 Then 'use assaylevel
                            str1 = "CHARHELPER1"
                        ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                            str1 = "NOMCONC"
                        ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                            str1 = "CHARHELPER1"
                        Else
                            str1 = "CHARHELPER1"
                        End If

                    End If

                    tblH1 = dv2.ToTable("H1", True, str1)
                    numH1 = tblH1.Rows.Count


                    'reset dv2 with a different sort
                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug

                    'things we need to know:
                    '-number of RunIDs
                    '   -number of Run Identifiers within each Run ID
                    '   -number of QC Levels within each Run ID
                    Dim tbl2a As System.Data.DataTable
                 
                    strS = "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    dv1 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    tbl2a = dv1.ToTable

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

                    'find number of different Run Identifiers
                    tblRID = dv2.ToTable("b", True, "RUNID")
                    numRID = tblRID.Rows.Count

                    'redo to include different charhelper2's
                    Dim numRID2 As Short
                    tblRID = dv2.ToTable("b", True, "RUNID", "CHARHELPER2")
                    numRID2 = tblRID.Rows.Count

                    'establish table of level numbers
                    'must be sorted by nomconc!
                    'make new dv
                    Dim dvNL As New DataView(tbl2, strF2, "NOMCONC ASC", DataViewRowState.CurrentRows)
                    'tblLevels = dvNL.ToTable("b", True, "NOMCONC")
                    'intNumLevels = tblLevels.Rows.Count

                    '20181127 LEE: Account for Stock (8) or Spiking (9) solution
                    If boolBSNR Then
                        tblLevels = dvNL.ToTable("b", True, "NOMCONC")
                    Else
                        'redo tblLevels to include different charhelper1's
                        'tblLevels = dvNL.ToTable("b", True, "NOMCONC", "CHARHELPER1")
                        '20180823 LEE:
                        If INTQCLEVELGROUP = 0 Then 'use assaylevel
                            tblLevels = dvNL.ToTable("b", True, "NOMCONC", "CHARHELPER1")
                        ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                            tblLevels = dvNL.ToTable("b", True, "NOMCONC")
                        ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                            tblLevels = dvNL.ToTable("b", True, "NOMCONC", "CHARHELPER1")
                        Else
                            tblLevels = dvNL.ToTable("b", True, "NOMCONC", "CHARHELPER1")
                        End If
                    End If

                    intNumLevels = tblLevels.Rows.Count

                    '20180823 LEE:
                    'need to do this to get level labels
                    Dim tblLabels As DataTable = dvNL.ToTable("b", True, "NOMCONC", "CHARHELPER1")

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
                        If INTQCLEVELGROUP = 1 Or boolBSNR Then 'use NomConc

                        Else
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
                        End If

                    Next

                    'establish combo table of level numbers and run identifiers
                    Dim tblLevelsRID As System.Data.DataTable
                    Dim intNumLevelsRID As Short

                    '20180823 LEE:
                    If INTQCLEVELGROUP = 1 Or boolBSNR Then 'use NomConc
                        tblLevelsRID = dv2.ToTable("c", True, "NOMCONC", "CHARHELPER2")
                    Else
                        tblLevelsRID = dv2.ToTable("c", True, "NOMCONC", "CHARHELPER1", "CHARHELPER2")
                    End If
                    intNumLevelsRID = tblLevelsRID.Rows.Count

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

                    '1=level1 normal stats, 2 = level2 normal stats, 3=level1 2nd stats, 4 = level2 2nd stats, 
                    '5=level1 normal stats fullprec, 6 = level2 normal stats fullprec, 7=level1 2nd stats fullprec, 8 = level2 2nd stats fullprec

                    'generate table
                    intTblRows = 0
                    intTblRows = intTblRows + 3 'for header
                    intTblRows = intTblRows + 1 'for blank row

                    If boolAdHocStabCompColumns Then
                        intTblRows = intTblRows + (intRowsX) 'for number of data rows
                        intTblRows = intTblRows + (1)   'for one blank rows after each run set
                        intTblRows = intTblRows + (3)   'for one %Diff rows after each run set
                    Else
                        intTblRows = intTblRows + (intRowsX * numH2) 'for number of data rows
                        intTblRows = intTblRows + (1 * numH2)   'for one blank rows after each run set
                        intTblRows = intTblRows + (3 * numH2)   'for one %Diff rows after each run set
                    End If


                    Dim numSets As Short
                    numSets = numRID

                    If numRID = numRID2 Then
                        numSets = 1
                    Else
                        If numH2 = 1 And numRID = 1 Then
                            numSets = numRID
                        ElseIf numH2 > 1 And numRID = 1 Then
                            numSets = numH2
                        End If
                    End If

                    numSets = numH2 'I don't understand the previous logic
                    '20180419 LEE: New logic. numSets is the same thing as numH2

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()

                    'ctExp is for deleting rows if no outliers occur

                    If boolAdHocStabCompColumns Then
                        intTblRows = intTblRows + intCSN
                    Else
                        intTblRows = intTblRows + (intCSN * intNumRuns) + (intCSN * numSets)
                        If intCSN > 0 Then
                            intTblRows = intTblRows + (2 * intNumRuns) - 3 'for two blank rows after each Mean/Bias/n set, except last set
                        End If
                        intTblRows = intTblRows + (2 * numH2) 'for two blank rows after each run identifier Mean/Bias/n set
                    End If

                    If boolQCREPORTACCVALUES Or boolAdHocStabCompColumns Then

                    Else
                        'For Count2 = 1 To intNumRuns 'must loop because this table may have several sections
                        'intTblRows = intTblRows + (3 * intNumRuns) 'for stats headings
                        '20180419 LEE:
                        intTblRows = intTblRows + (intNumRuns) 'for stats headings
                        ctExp = ctExp + 2 'for stats headings

                        'Increment for Statistics Sections
                        intTblRows = intTblRows + (intCSN * intNumRuns) + (intCSN * numH2)
                        ctExp = ctExp + intCSN
                        '20180419 LEE:

                        'rows for %Diff
                        intTblRows = intTblRows + (numH2 * 2)
                        ctExp = ctExp + 2

                        If intCSN > 0 Then
                            intTblRows = intTblRows + (1 * intNumRuns) - 1 'for a blank row after each Mean/Bias/n set, except last set

                            intTblRows = intTblRows + (2 * numSets) 'for two blank rows after each run identifier Mean/Bias/n set

                            intTblRows = intTblRows + 4 'add additional rows for stability in second stats section

                        End If

                    End If

                    If intNumRuns = 1 Then
                    Else
                        intTblRows = intTblRows + 2 'add additional rows for stability
                        ctExp = ctExp + 2
                    End If

                    wrdSelection = wd.Selection()

                    Dim intCols As Short
                    Dim intDiffInc As Short 'column increment if diff stats column is true

                    'boolSTATSDIFFCOL if for Accuracy column
                    If boolAdHocStabCompColumns Then
                        'ignore numLevels and make columns based on CHARHELPER2
                        intNumLevels = 1
                        If boolSTATSDIFFCOL Then
                            intCols = (numH2 * 2) + 1
                            intDiffInc = 2
                        Else
                            intCols = numH2 + 1
                            intDiffInc = 1
                        End If
                    Else
                        If boolSTATSDIFFCOL Then
                            intCols = (intNumLevels * 2) + 1
                            intDiffInc = 2
                        Else
                            intCols = intNumLevels + 1
                            intDiffInc = 1
                        End If
                        If boolColDiff And numSets > 1 Then
                            intCols = intCols + (intNumLevels)
                            intDiffInc = intDiffInc + 1
                        End If
                        If boolSampleNameCol Then
                            intCols = intCols + 1
                        End If
                    End If

                    If boolAdHocStabCompColumns Then
                        ReDim arrS(10, intCols)
                    Else
                        ReDim arrS(10, intNumLevels * numH2)
                    End If

                    Dim tbl As Word.Table

                    '20190225 LEE is Int Std
                    'New feature allows user to choose IS in Adv Table Config
                    'at this point, assign this to boolX
                    If boolRCPA = False And boolRCConc = False And boolRCPARatio = False Then
                        boolX = True
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
                        tbl = .Selection.Tables.Item(1)

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                        '.Selection.Tables.Item(1).Columns.Item(1).Width = 86

                        .Selection.Tables.Item(1).Select()

                        'remove border, but leave top and bottom
                        Call removeBorderButLeaveTopAndBottom(wd)

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

                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        '20180309 LEE: Do not bottom border
                        '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        '20180313 LEE:
                        Dim intCol1 As Short
                        Dim intCol2 As Short
                        Dim intColSN As Short
                        Dim intColDiff As Short

                        If boolSampleNameCol Then
                            intCol1 = 1
                            intColSN = 2
                            intCol2 = 3
                            intColDiff = 4
                        Else
                            intCol1 = 1
                            intCol2 = 2
                            intColSN = intCol1
                            intColDiff = 3
                        End If

                        'Enter QC ID row titles
                        .Selection.Tables.Item(1).Cell(intCRow, intCol2).Select()

                        'get labels
                        Dim tblL = tblTableProperties
                        strF = "ID_TBLREPORTTABLE = " & idTR
                        Dim rowsL() As DataRow = tblL.Select(strF)
                        Dim strFirst As String

                        If boolAdHocStabCompColumns Then
                            int8 = int8 + 1
                        End If

                        If rowsL.Length = 0 Then
                            strFirst = ""
                            strDiff = "%Difference"
                        Else
                            str1 = ""
                            If boolMEANACCURACY Then
                                str1 = "Mean Accuracy:"
                            ElseIf BOOLDIFFERENCE Then
                                str1 = "%Difference:"
                            ElseIf boolRECOVERY Then
                                str1 = "%Recovery:"
                            End If
                            '20190225 LEE:
                            strDiff = Mid(str1, 1, Len(str1) - 1)
                            strFirst = NZ(rowsL(0).Item("CHARTITLELEG"), str1)
                            'if there is an '=' sign, then remove it
                            strFirst = Replace(strFirst, " = ", "", 1, -1, CompareMethod.Text)
                            strFirst = Replace(strFirst, "= ", "", 1, -1, CompareMethod.Text)
                            strFirst = Replace(strFirst, " =", "", 1, -1, CompareMethod.Text)
                        End If

                        For Count2 = 0 To intNumLevels - 1

                            'If numRID > 1 Then 'use run identifier instead of QC
                            '    var2 = tblLevels.Rows.Item(Count2).Item("CHARHELPER2")
                            'Else
                            '    var2 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                            'End If
                            var1 = tblLevels.Rows.Item(Count2).Item("NOMCONC")

                            If INTQCLEVELGROUP = 1 Or boolBSNR Then 'use NomConc 
                                var2 = ""
                                'must get label from somewhere else
                                Dim rowsLabels() As DataRow = tblLabels.Select("NOMCONC = " & CSng(var1))
                                If rowsLabels.Length = 0 Then
                                    var2 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                                Else
                                    var2 = rowsLabels(0).Item("CHARHELPER1")
                                End If
                            Else
                                var2 = tblLevels.Rows.Item(Count2).Item("CHARHELPER1")
                            End If

                            '20181127 LEE:
                            var2 = NZ(var2, "")

                            If Len(var2) = 0 Then
                                If BOOLINCLUDEDATE Then
                                Else
                                    intCRow = intCRow - 1
                                End If
                                Exit For
                            End If
                            'var3 = var2

                            '20181127 LEE:
                            If boolBSNR Then
                                strCombTitle1 = ""
                            Else
                                var3 = ReturnStdQC(var2.ToString)
                                If boolAdHocStabCompColumns Then
                                    strCombTitle1 = var3
                                Else
                                    .Selection.TypeText(Text:=var3)
                                End If
                            End If


                            '******determine if the level is a diln level
                            Dim strE As String
                            'var3 = var2 ' & ChrW(10) & var1 & " " & strConcUnits
                            ''strE = ChrW(10) & var1 & " " & strConcUnits
                            '.Selection.TypeText(Text:=var3)

                            dv2.RowFilter = ""
                            strF = strF2 & " AND NOMCONC = " & CDbl(var1)
                            dv2.RowFilter = strF
                            'check for aliquot factor
                            Dim numDS As Single
                            If dv2.Count = 0 Then

                            Else
                                numDS = NZ(dv2(0).Item("ALIQUOTFACTOR"), 1)
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
                                        'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(CDbl(Val(var1)), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                                        arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & Format(CDbl(Val(var1)), GetRegrDecStr(LSigFig)) & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
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

                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            'boolcoldif
                            If boolSTATSDIFFCOL Or (boolColDiff And numSets > 1) Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If

                        Next Count2

                        intCRow = intCRow + 1

                        If boolSampleNameCol Then
                            'enter sample name here
                            .Selection.Tables.Item(1).Cell(intCRow, intColSN).Select()
                            .Selection.TypeText(Text:=strSampleNameCol)
                        End If

                        'Enter nom. conc. row titles
                        .Selection.Tables.Item(1).Cell(intCRow, intCol2).Select()


                        For Count2 = 0 To intNumLevels - 1

                            'var1 = arrBCQCs(3, Count2)
                            If boolLUseSigFigs Then
                                var1 = SigFigOrDecString(NZ(tblLevels.Rows.Item(Count2).Item("NOMCONC"), 0), LSigFig, False)
                            Else
                                var1 = RoundToDecimalRAFZ(NZ(tblLevels.Rows.Item(Count2).Item("NOMCONC"), 0), LSigFig)
                            End If

                            If boolX And Len(strISConc) <> 0 Then
                                If LboolNomConcParen Then
                                    var3 = "(" & strISConc & ")"
                                Else
                                    var3 = strISConc
                                End If
                            Else
                                If boolLUseSigFigs Then
                                    If LboolNomConcParen Then
                                        var3 = "(" & DisplayNum(var1, LSigFig, False) & ChrW(160) & strConcUnits & ")"
                                    Else
                                        var3 = DisplayNum(var1, LSigFig, False) & ChrW(160) & strConcUnits
                                    End If
                                Else
                                    If LboolNomConcParen Then
                                        var3 = "(" & Format(var1, GetRegrDecStr(LSigFig)) & ChrW(160) & strConcUnits & ")"
                                    Else
                                        var3 = Format(var1, GetRegrDecStr(LSigFig)) & ChrW(160) & strConcUnits
                                    End If
                                End If
                            End If

                            If IsNumeric(var3) Then
                                If boolLUseSigFigs Then
                                    str1 = DisplayNum(var3, LSigFig, False)
                                    '.Selection.TypeText(Text:=DisplayNum(var3, LSigFig, False))
                                Else
                                    str1 = Format(var3, GetRegrDecStr(LSigFig))
                                    '.Selection.TypeText(Text:=Format(var3, GetRegrDecStr(LSigFig)))
                                End If
                            Else
                                str1 = var3
                                '.Selection.TypeText(Text:=var3)
                            End If

                            'may need to add additional information
                            'don't add Peak Area, etc
                            'labeled later
                            If boolRCPA Then
                                var4 = str1 ' & ChrW(10) & "(Peak Area)"
                            ElseIf boolRCPARatio Then
                                var4 = str1 ' & ChrW(10) & "(Peak Area Ratio)"
                            Else
                                var4 = str1
                            End If

                            If boolAdHocStabCompColumns Then
                                strCombTitle2 = var4
                            Else
                                .Selection.TypeText(Text:=var4)
                            End If

                            If boolAdHocStabCompColumns Then
                                If Len(strCombTitle1) = 0 Then
                                    strCombTitle = strCombTitle2
                                Else
                                    strCombTitle = strCombTitle1 & " " & strCombTitle2
                                End If
                                .Selection.Tables.Item(1).Cell(intCRow - 2, intCol2).Select()

                                '20181127 LEE
                                If boolBSNR Then
                                Else
                                    .Selection.TypeText(Text:=strCombTitle)
                                End If

                                .Selection.Tables.Item(1).Cell(intCRow, intCol2).Select()
                            End If


                            'If boolNONELEG Then
                            '    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            'Else
                            '    If boolSTATSDIFFCOL Then
                            '        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            '        .Selection.TypeText(Text:=ReturnDiffLabel)
                            '        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            '    End If

                            '    'If boolColDiff And numSets > 1 Then
                            '    If boolColDiff And numSets > 1 Then
                            '        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            '        .Selection.TypeText(Text:=strFirst)
                            '    End If
                            'End If

                            If boolSTATSDIFFCOL Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                .Selection.TypeText(Text:=ReturnDiffLabel)
                                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If

                            'If boolColDiff And numSets > 1 Then
                            If boolColDiff And numSets > 1 Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                .Selection.TypeText(Text:=strFirst)
                            End If

                            If Count2 = intNumLevels Then
                            Else
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If

                        Next Count2

                        .Selection.Tables.Item(1).Cell(intCRow, intCol1).Select()
                        'bottom border this row
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                       
                        'enter Watson Run ID column header
                        .Selection.Tables.Item(1).Cell(intCRow, intCol1).Select()

                        If BOOLINCLUDEDATE Then
                            'enter 'Label' header
                            .Selection.Tables.Item(1).Cell(intCRow - 2, intCol1).Select()
                            .Selection.TypeText(strLabel)
                            .Selection.Tables.Item(1).Cell(intCRow - 1, intCol1).Select()
                            .Selection.TypeText(strWRunId)
                            .Selection.Tables.Item(1).Cell(intCRow, intCol1).Select()
                            '.Selection.TypeText("(Analysis Date)")
                            '20180420 LEE:
                            .Selection.TypeText("(" & GetAnalysisDateLabel(intTableID) & ")")
                        Else
                            'enter 'Label' header
                            .Selection.Tables.Item(1).Cell(intCRow - 1, intCol1).Select()
                            .Selection.TypeText(strLabel)
                            .Selection.Tables.Item(1).Cell(intCRow, intCol1).Select()
                            .Selection.TypeText(strWRunId)
                        End If

                        'begin entering data'
                        intCRow = intCRow + 2
                        intStart = intCRow ' 5
                        int1 = intCRow ' 5 'row position counter
                        Dim boolExit As Boolean = False
                        Dim numArrs As Short

                        'If numH2 = 1 And numRID = 1 Then
                        '    numSets = numRID
                        'ElseIf numH2 > 1 And numRID = 1 Then
                        '    numSets = numH2
                        'End If

                        If numRID = numRID2 Then
                            numSets = 1
                        Else
                            If numH2 = 1 And numRID = 1 Then
                                numSets = numRID
                            ElseIf numH2 > 1 And numRID = 1 Then
                                numSets = numH2
                            End If
                        End If

                        numSets = numH2 'I don't understand the previous logic
                        '20180419 LEE: Logic has changed
                        'numH2 is the same as numSets

                        Dim intoStart As Short
                        intoStart = intStart + 1 'account for strH2 entry

                        'For Count6 = 0 To intNumLevels - 1

                        varNom = tblLevels.Rows.Item(Count6).Item("NOMCONC")

                        'now do each set
                        Dim intCol As Short = 1
                        Dim intColD As Short = 1

                        If boolSampleNameCol Then
                            intCol = intCol1 ' intColSN
                        Else
                            intCol = intCol1
                        End If

                        Dim numMaxRows As Short = 6 'default
                        Dim strH2 As String
                        Dim rowsRuns() As DataRow
                        Dim strFR As String

                        If boolAdHocStabCompColumns Then
                            If numH2 = 1 Then
                                numMaxRows = intRowsX
                            Else
                                numMaxRows = 0
                                For Count2 = 0 To numH2 - 1 '20190117 LEE: numH2 is numRuns

                                    strH2 = NZ(tblH2.Rows(Count2).Item("CHARHELPER2"), "[NONE]")

                                    If StrComp(strH2, "[NONE]", CompareMethod.Text) = 0 Then
                                        strFR = "CHARHELPER2 = ''"
                                        rowsRuns = tblRID.Select(strFR, "RUNID ASC")
                                        intNumRuns = rowsRuns.Length
                                        If intNumRuns = 0 Then
                                            'check for null
                                            strFR = "CHARHELPER2 IS NULL"
                                            rowsRuns = tblRID.Select(strFR, "RUNID ASC")
                                            intNumRuns = rowsRuns.Length
                                        End If
                                    Else
                                        strFR = "CHARHELPER2 = '" & strH2 & "'"
                                        rowsRuns = tblRID.Select(strFR, "RUNID ASC")
                                        intNumRuns = rowsRuns.Length
                                    End If

                                    For Count5 = 0 To intNumRuns - 1

                                        'varNom = tblLevels.Rows.Item(Count5).Item("NOMCONC")
                                        var10 = rowsRuns(Count5).Item("RUNID")
                                        'start evaluating data
                                        dv2.RowFilter = ""
                                        'don't know why, but must make a long filter here or
                                        'both analytes get returned in dv2.rowfilter

                                        If StrComp(strH2, "[NONE]", CompareMethod.Text) = 0 Then
                                            strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND (CHARHELPER2 IS NULL OR CHARHELPER2 = '')"
                                        Else
                                            strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND CHARHELPER2 = '" & strH2 & "'"
                                        End If

                                        dv2.RowFilter = strF
                                        int2 = dv2.Count

                                        If int2 > numMaxRows Then
                                            numMaxRows = int2
                                        End If

                                    Next Count5

                                Next Count2

                            End If

                        End If

                        For Count2 = 0 To numH2 - 1

                            If boolAdHocStabCompColumns Then
                                intStart = intCRow ' 5
                                If boolSTATSDIFFCOL Then
                                    If Count2 = 0 Then
                                        intCol = intCol + 1
                                    Else
                                        intCol = intCol + 2
                                    End If
                                Else
                                    intCol = intCol + 1
                                End If
                            Else

                            End If


                            strH2 = NZ(tblH2.Rows(Count2).Item("CHARHELPER2"), "[NONE]")

                            'enter strH2
                            .Selection.Tables.Item(1).Cell(intStart, intCol).Select()
                            If boolAdHocStabCompColumns Then
                                'put strH2 at column head
                                .Selection.Tables.Item(1).Cell(intStart - 2, intCol).Range.Text = strH2

                                If boolSTATSDIFFCOL Then
                                    .Selection.Tables.Item(1).Cell(intStart - 2, intCol + 1).Range.Text = ReturnDiffLabel()
                                End If

                            Else
                                If StrComp(strH2, "[NONE]", CompareMethod.Text) = 0 Then
                                Else
                                    .Selection.TypeText(strH2)
                                    'intStart = intStart + 1
                                End If
                                'intoStart = intStart
                            End If
                            intoStart = intStart

                            If StrComp(strH2, "[NONE]", CompareMethod.Text) = 0 Then
                                strFR = "CHARHELPER2 = ''"
                                rowsRuns = tblRID.Select(strFR, "RUNID ASC")
                                intNumRuns = rowsRuns.Length
                                If intNumRuns = 0 Then
                                    'check for null
                                    strFR = "CHARHELPER2 IS NULL"
                                    rowsRuns = tblRID.Select(strFR, "RUNID ASC")
                                    intNumRuns = rowsRuns.Length
                                End If
                            Else
                                strFR = "CHARHELPER2 = '" & strH2 & "'"
                                rowsRuns = tblRID.Select(strFR, "RUNID ASC")
                                intNumRuns = rowsRuns.Length
                            End If



                            'must determine if there are any outliers
                            intExp = 0
                            int12 = -1
                            For Count5 = 0 To intNumRuns - 1

                                For Count3 = 0 To intNumLevels - 1

                                    int12 = int12 + 1
                                    varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")

                                    'varNom = tblLevels.Rows.Item(Count5).Item("NOMCONC")
                                    var10 = rowsRuns(Count5).Item("RUNID")
                                    'start evaluating data
                                    dv2.RowFilter = ""
                                    'don't know why, but must make a long filter here or
                                    'both analytes get returned in dv2.rowfilter

                                    If StrComp(strH2, "[NONE]", CompareMethod.Text) = 0 Then
                                        strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND (CHARHELPER2 IS NULL OR CHARHELPER2 = '')"
                                    Else
                                        strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND CHARHELPER2 = '" & strH2 & "'"
                                    End If

                                    dv2.RowFilter = strF
                                    int2 = dv2.Count

                                    'create rows1 from tbl1 which will contain data
                                    'strF = ""

                                    Erase rows2

                                    Dim tbl2S As System.Data.DataTable = dv2.ToTable
                                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

                                    rows2 = tbl2S.Select(strF, strS)
                                    int3 = rows2.Length
                                    intRowsX = int3 'new introwsx

                                    Dim CountA As Short
                                    Dim CountB As Short

                                    'evaluate data
                                    'For Count4 = 0 To intRowsX - 1 'int3 - 1
                                    For Count4 = 0 To rows2.Length - 1 'int3 - 1
                                        var1 = NZ(rows2(Count4).Item("ELIMINATEDFLAG"), "N")
                                        If StrComp(var1, "Y", CompareMethod.Text) = 0 Then
                                            intExp = intExp + 1
                                            boolHasOutlier = True
                                            boolExit = True
                                            Exit For
                                        End If
                                    Next

                                    If boolExit Then
                                        Exit For
                                    End If

                                Next Count3

                            Next Count5

                            'For Count10 = 0 To numRID - 1

                            Dim boolEnterDiff As Boolean
                            Dim maxRows As Short = 0

                            Dim arrMaxRows(10)

                            For Count5 = 0 To 10
                                arrMaxRows(Count5) = 0
                            Next


                            int12 = -1

                            'start filling in data by columns
                            'intRowsX = 0
                            'For Count3 = 0 To (intNumLevels * int11) - 1 Step int11
                            For Count3 = 0 To intNumLevels - 1

                                int12 = int12 + 1
                                varNom = tblLevels.Rows.Item(int12).Item("NOMCONC")

                                If boolAdHocStabCompColumns Then

                                Else
                                    intStart = intoStart
                                End If

                                nI = 0

                                'start entering data
                                dv2.RowFilter = ""
                                'don't know why, but must make a long filter here or
                                'both analytes get returned in dv2.rowfilter
                                If StrComp(strH2, "[NONE]", CompareMethod.Text) = 0 Then
                                    strF = strF2 & " AND NOMCONC = " & varNom & " AND (CHARHELPER2 IS NULL OR CHARHELPER2 = '')"
                                Else
                                    strF = strF2 & " AND NOMCONC = " & varNom & " AND CHARHELPER2 = '" & strH2 & "'"
                                End If
                                'strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom

                                dv2.RowFilter = strF
                                int2 = dv2.Count

                                nI = nI + int2

                                'create rows1 from tbl1 which will contain data
                                strF = ""

                                Erase rows2

                                Dim tbl2se As System.Data.DataTable = dv2.ToTable
                                'now do rows2E
                                strF = ""
                                Erase rows2E
                                'rows2E = tbl1.Select(strF)
                                If gAllowExclSamples And LAllowExclSamples Then
                                    rows2E = tbl2se.Select("(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL) AND BOOLEXCLSAMPLE = 0")
                                Else
                                    rows2E = tbl2se.Select("ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL")
                                End If
                                nE = rows2E.Length

                                'now do rows2A
                                Erase rows2A
                                If gAllowExclSamples And LAllowExclSamples Then
                                    rows2A = tbl2se.Select("RUNID > 0")
                                Else
                                    rows2A = tbl2se.Select("RUNID > 0")
                                End If

                                ReDim arrMaxRows(intNumLevels)
                                'set arrmaxrows

                                maxRows = 0

                                For Count10 = 0 To intNumRuns - 1

                                    boolOutHeadI = False
                                    boolOutHeadE = False

                                    'strM = "Creating Summary of " & strTempInfo & " Final Extract Stability Table For " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & "..."
                                    frmH.lblProgress.Text = strM1 & ChrW(10) & "Processing Run ID " & var10
                                    frmH.Refresh()

                                    Dim strH As String

                                    'enter runid
                                    'var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")
                                    var10 = rowsRuns(Count10).Item("RUNID")

                                    'If numRID > 1 And intNumRuns > 1 Then
                                    '    strH = NZ(tblRID.Rows(Count2).Item("CHARHELPER2"), "[NONE]")
                                    'Else
                                    '    strH = NZ(tblRID.Rows(Count10).Item("CHARHELPER2"), "[NONE]")
                                    'End If
                                    'previous logic:  wtf???
                                    strH = strH2

                                    int1 = intStart
                                    If Count3 = 0 Then

                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(intStart, intCol1).Select()
                                        Else
                                            If StrComp(strH2, "[NONE]", CompareMethod.Text) = 0 Then
                                                .Selection.Tables.Item(1).Cell(intStart, intCol1).Select()
                                            Else
                                                .Selection.Tables.Item(1).Cell(intStart + 1, intCol1).Select()
                                                'intStart = intStart + 1
                                            End If
                                        End If

                                        .Selection.TypeText(CStr(var10))
                                        .Selection.Tables.Item(1).Cell(intStart, intCol1).Select()
                                        If StrComp(strH, "[NONE]", CompareMethod.Text) = 0 Then
                                        Else
                                            'int1 = int1 + 1
                                            '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                                            '.Selection.TypeText(strH)
                                            '.Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                                        End If
                                        If BOOLINCLUDEDATE Then
                                            'int1 = int1 + 1
                                            If boolAdHocStabCompColumns Then
                                                .Selection.Tables.Item(1).Cell(intStart + 1, intCol1).Select()
                                            Else
                                                .Selection.Tables.Item(1).Cell(intStart + 2, intCol1).Select()
                                            End If

                                            str1 = GetDateFromRunID(NZ(var10, 0), LDateFormat, intGroup, idTR)
                                            .Selection.TypeText("(" & str1 & ")")
                                            .Selection.Tables.Item(1).Cell(intStart, intCol1).Select()
                                        End If
                                        ' .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                    End If

                                    If boolAdHocStabCompColumns Then
                                        .Selection.Tables.Item(1).Cell(intStart, intCol).Select()
                                    Else
                                        If Count3 = 0 Then
                                            .Selection.Tables.Item(1).Cell(intStart, intCol1 + Count3 + intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(intStart, intCol1 + (Count3 * intDiffInc) + intCol).Select()
                                        End If
                                    End If

                                    '*******

                                    'determine hi and lo (nom*flagpercent)
                                    strF = "CONCENTRATION = '" & varNom & "'"
                                    'rows10 = tblBCQCs.Select(strF)

                                    If boolRCPA Or boolRCPARatio Then
                                        hi = 0
                                        lo = 0
                                        arrFP(1, int12) = 15
                                        arrFP(2, int12) = 15
                                        v1 = 15
                                        v2 = 15
                                        vU = 0
                                    Else

                                        'determine hi and lo (nom*flagpercent)
                                        'If Len(NZ(vAnalyteID, "")) = 0 Then
                                        '    strF = "CONCENTRATION = " & varNom & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10
                                        'Else
                                        '    strF = "CONCENTRATION = " & varNom & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & varNom & " AND RUNID = " & var10
                                        'End If
                                        'if Conc < 1, then the query return 0 records
                                        'must do something different
                                        var1 = GetANALYTEFLAGPERCENT(varNom, var10, vAnalyteID)

                                        'var1 = CDec(NZ(rows10(0).Item("FLAGPERCENT"), 15))
                                        arrFP(1, int12) = var1
                                        arrFP(2, int12) = var1
                                        Call SetHighAndLowCriteria(varNom, var1, var1, hi, lo)
                                        v1 = var1
                                        v2 = var1
                                        vU = 0
                                    End If

                                    'start entering data
                                    dv2.RowFilter = ""
                                    'don't know why, but must make a long filter here or
                                    'both analytes get returned in dv2.rowfilter
                                    If StrComp(strH, "[NONE]", CompareMethod.Text) = 0 Then
                                        strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND (CHARHELPER2 IS NULL OR CHARHELPER2 = '')"
                                    Else
                                        strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND CHARHELPER2 = '" & strH & "'"
                                    End If
                                    'strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom

                                    dv2.RowFilter = strF
                                    int2 = dv2.Count
                                    If int2 = 0 Then 'skip
                                    Else

                                        nI = nI + int2

                                        'create rows1 from tbl1 which will contain data
                                        strF = ""

                                        Erase rows2

                                        Dim tbl2s As System.Data.DataTable = dv2.ToTable
                                        rows2 = tbl2s.Select("RUNID > -1")

                                        If boolAdHocStabCompColumns Then
                                            maxRows = intRowsX
                                        Else
                                            int3 = rows2.Length
                                            intRowsX = int3 'new introwsx
                                            If int3 > arrMaxRows(Count3) Then
                                                arrMaxRows(Count3) = int3
                                                maxRows = int3
                                            End If
                                        End If


                                        'REDO HI/LO
                                        vU = rows2(0).Item("BOOLUSEGUWUACCCRIT")
                                        If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                            v1 = CDec(NZ(rows2(0).Item("NUMMAXACCCRIT"), 0))
                                            v2 = CDec(NZ(rows2(0).Item("NUMMINACCCRIT"), 0))
                                            arrFP(1, int12) = v1
                                            arrFP(2, int12) = v2
                                            Call SetHighAndLowCriteria(varNom, v1, v2, hi, lo)

                                        End If

                                        'now do rows2E
                                        strF = ""
                                        'Erase rows2E
                                        ''rows2E = tbl1.Select(strF)
                                        'If gAllowExclSamples And LAllowExclSamples Then
                                        '    rows2E = tbl2s.Select("(ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL) AND BOOLEXCLSAMPLE = 0")
                                        'Else
                                        '    rows2E = tbl2s.Select("ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL")
                                        'End If

                                        'nE = rows2E.Length

                                        For Count4 = 0 To maxRows - 1 ' intRowsX - 1 'int3 - 1

                                            boolEnterDiff = False
                                            boolOC = False

                                            If boolAdHocStabCompColumns Then
                                                .Selection.Tables.Item(1).Cell(intStart + Count4, intCol).Select()
                                            Else
                                                If Count3 = 0 Then
                                                    .Selection.Tables.Item(1).Cell(intStart + Count4, intCol1 + Count3 + intColSN).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(intStart + Count4, intCol1 + (Count3 * intDiffInc) + intColSN).Select()
                                                End If
                                            End If


                                            If Count4 > intRowsX - 1 Then

                                                If boolQCNA Then
                                                    str1 = "NA"
                                                Else
                                                    str1 = ""
                                                End If

                                                .Selection.TypeText(str1)
                                                boolEnterDiff = False
                                                boolOC = True

                                                If boolSampleNameCol Then
                                                    .Selection.Tables.Item(1).Cell(intStart + Count4, intColSN).Range.Text = "NA"
                                                End If

                                            Else
                                                var1 = rows2(Count4).Item("CONCENTRATION")
                                                varConc = var1
                                                varAnal = NZ(rows2(Count4).Item("ANALYTEAREA"), 0)
                                                varIS = NZ(rows2(Count4).Item("INTERNALSTANDARDAREA"), 0)
                                                var1 = NZ(var1, 0)
                                                numDF = rows2(Count4).Item("ALIQUOTFACTOR")
                                                var1 = var1 / numDF
                                                If boolRCConc Then
                                                    If boolLUseSigFigs Then
                                                        var2 = SigFigOrDec(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                    Else
                                                        var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If
                                                Else

                                                    '*****
                                                    If boolX Then
                                                        If boolLUseSigFigsArea Then
                                                            var1 = SigFigArea(RoundToDecimalA(varIS / numDF, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                        Else
                                                            var1 = RoundToDecimalRAFZ(varIS / numDF, LSigFigArea)
                                                        End If
                                                        var2 = var1
                                                    Else
                                                        If boolRCPARatio Then
                                                            If boolLUseSigFigsAreaRatio Then
                                                                var2 = SigFigAreaRatio(RoundToDecimalA((varAnal / numDF) / varIS, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                            Else
                                                                var2 = RoundToDecimalRAFZ((varAnal / numDF) / varIS, LSigFigAreaRatio)
                                                            End If
                                                        Else
                                                            If boolLUseSigFigsArea Then
                                                                var1 = SigFigArea(RoundToDecimalA(varAnal, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                            Else
                                                                var1 = RoundToDecimalRAFZ(varAnal, LSigFigArea)
                                                            End If
                                                            var2 = var1
                                                        End If
                                                    End If
                                                    '*****

                                                End If

                                                var1 = NZ(rows2(Count4).Item("ELIMINATEDFLAG"), "N")
                                                var3 = NZ(rows2(Count4).Item("BOOLEXCLSAMPLE"), 0)

                                                If IsDBNull(varConc) And boolRCConc Then

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
                                                    boolHasOutlier = True
                                                    intLeg = intLeg + 1
                                                    strA = ChrW(intLeg + intLegStart)

                                                    '20160305 LEE:
                                                    'Added DECISIONREASON code
                                                    Dim var6
                                                    'Remember, tblAssignedSamples does not have DECISIONREASON
                                                    var6 = "No Value: " & GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))
                                                    'Set Legend String
                                                    '20190225 LEE:'20190225 LEE: This is temporary until parentheses are implemented throughout at a later date
                                                    If Count2 <> 0 Then
                                                        strDiff = ""
                                                    End If
                                                    If boolAdHocStabCompColumns Then
                                                        str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, True, strDiff)
                                                    Else
                                                        'str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, False)
                                                        '20190225' %Diff will have parentheses, soon stats will
                                                        str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, False, strDiff)
                                                    End If

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
                                                    boolHasOutlier = True
                                                    intLeg = intLeg + 1
                                                    strA = ChrW(intLeg + intLegStart)

                                                    '20160305 LEE:
                                                    'Added DECISIONREASON code
                                                    Dim var6
                                                    'Remember, tblAssignedSamples does not have DECISIONREASON
                                                    var6 = GetDECISIONREASONValue(boolExFromAS, vAnalyteID, rows2(Count4))
                                                    'Set Legend String
                                                    '20190225 LEE: This is temporary until parentheses are implemented throughout at a later date
                                                    If Count2 <> 0 Then
                                                        strDiff = ""
                                                    End If
                                                    If boolAdHocStabCompColumns Then
                                                        str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, True, strDiff)
                                                    Else
                                                        'str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, False)
                                                        '20190225' %Diff will have parentheses
                                                        str1 = GetLegendStringExcluded(v1, v2, vU, var6, intTableID, False, strDiff)
                                                    End If

                                                    'Add to Legend Array
                                                    ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                    If boolRedBoldFont Then
                                                        .Selection.Font.Bold = True
                                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                    End If

                                                    If boolRCConc Then
                                                        If boolLUseSigFigs Then
                                                            .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                        Else
                                                            .Selection.TypeText(Text:=CStr(Format(var2, GetRegrDecStr(LSigFig))))
                                                        End If

                                                    Else

                                                        '*****
                                                        var1 = var2
                                                        If boolRCPARatio Then
                                                            If boolLUseSigFigsAreaRatio Then
                                                                var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                            Else
                                                                var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                                            End If
                                                        Else
                                                            If boolLUseSigFigsArea Then
                                                                var1 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                            Else
                                                                var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                                            End If
                                                            var2 = var1
                                                        End If
                                                        '*****
                                                        .Selection.TypeText(Text:=CStr(var2))
                                                    End If
                                                    '.Selection.TypeText(Text:=CStr(var2))
                                                    Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                                Else

                                                    boolEnterDiff = True
                                                    'determine if value is outside acceptance criteria
                                                    'If (var2 > hi Or var2 < lo) And boolRCConc Then 'flag
                                                    If (OutsideAccCrit(var2, varNom, v1, v2, NZ(vU, 0))) And boolRCConc Then 'flag
                                                        intLeg = intLeg + 1
                                                        strA = ChrW(intLeg + intLegStart)

                                                        'Set Legend String
                                                        str1 = GetLegendStringIncluded(arrFP(1, int12), arrFP(2, int12), vU)
                                                        'Add to Legend
                                                        ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                        If boolRedBoldFont Then
                                                            .Selection.Font.Bold = True
                                                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                        End If

                                                        If boolRCConc Then
                                                            If boolLUseSigFigs Then
                                                                .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                            Else
                                                                .Selection.TypeText(Text:=CStr(Format(var2, GetRegrDecStr(LSigFig))))
                                                            End If
                                                        Else

                                                            var1 = var2
                                                            '*****
                                                            If boolRCPARatio Then
                                                                If boolLUseSigFigsAreaRatio Then
                                                                    var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                                Else
                                                                    var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                                                End If
                                                            Else
                                                                If boolLUseSigFigsArea Then
                                                                    var1 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                                Else
                                                                    var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                                                End If
                                                                var2 = var1
                                                            End If
                                                            '*****
                                                            .Selection.TypeText(Text:=CStr(var2))
                                                        End If

                                                        Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                    Else
                                                        If boolRCConc Then
                                                            If boolLUseSigFigs Then
                                                                .Selection.TypeText(Text:=CStr(DisplayNum(var2, LSigFig, False)))
                                                            Else
                                                                .Selection.TypeText(Text:=CStr(Format(var2, GetRegrDecStr(LSigFig))))
                                                            End If

                                                        Else

                                                            var1 = CDbl(var2)
                                                            '*****
                                                            If boolRCPARatio Then
                                                                If boolLUseSigFigsAreaRatio Then
                                                                    var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                                Else
                                                                    var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                                                End If
                                                            Else
                                                                If boolLUseSigFigsArea Then
                                                                    var1 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                                Else
                                                                    var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                                                End If
                                                                var2 = var1
                                                            End If
                                                            '*****
                                                            .Selection.TypeText(Text:=CStr(var2))
                                                        End If
                                                        boolEnterDiff = True
                                                    End If
                                                End If

                                                If boolSampleNameCol Then
                                                    Try
                                                        str1 = NZ(rows2(Count4).Item("SAMPLENAME"), "NA")
                                                    Catch ex As Exception
                                                        str1 = "NA"
                                                    End Try

                                                    .Selection.Tables.Item(1).Cell(intStart + Count4, intColSN).Range.Text = str1
                                                End If

                                            End If

                                            If boolSTATSDIFFCOL Then
                                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                                If boolEnterDiff Then
                                                    'var3 = Format(((var2 / varNom) - 1) * 100, strQCDec)
                                                    If boolTHEORETICAL Then
                                                        var3 = CalcREPercent(var2, varNom, intQCDec)
                                                        numTheor = 100 + CDec(var3)

                                                        'redo var3 as string in case %RE column is shown
                                                        var3 = Format(RoundToDecimal(CalcREPercent(var2, varNom, intQCDec), intQCDec), strQCDec)

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

                                    End If

                                    'increase row position counter
                                    'If Count2 = intNumRuns - 1 Then
                                    '    'int1 = int1 + intRowsX + 7
                                    '    intStart = int1 + int8 + 1
                                    'Else
                                    '    'int1 = int1 + intRowsX + 8
                                    '    intStart = int1 + int8 + 1
                                    'End If

                                    'intStart = intStart + intNumRuns\
                                    If Count10 = intNumRuns - 1 Then
                                    Else
                                        intStart = intStart + maxRows + 1
                                    End If


                                Next Count10


                                'here
                                '''''wdd.visible = True

                                If boolAdHocStabCompColumns Then
                                    intStart = intStart + numMaxRows + 1
                                Else
                                    'intStart = intStart + maxRows
                                    intStart = intStart + maxRows + 1
                                End If



                                'intStart = intStart + 1
                                .Selection.Tables.Item(1).Cell(intStart, intCol1).Select()

                                int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)

                                'begin doing stats
                                If boolQCREPORTACCVALUES Then
                                Else
                                    'herehere
                                    'intExp = 1 'must force in order to report stability
                                    'NO! This forces the "Summary Statistics Excluding Outlier Values" heading!
                                    If intExp = 0 Then
                                    Else
                                        '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                        '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                        'enter some blank spaces to fool PageBreak function
                                        '.selection.typetext(Text:="  ")
                                        If boolOutHeadE Or boolAdHocStabCompColumns Then
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

                                            .Selection.Tables.Item(1).Cell(int1 + 1, intCol1).Select()
                                        End If

                                    End If

                                    'Put it here instead
                                    'This isn't working right
                                    'intExp = 1 'must force in order to report stability

                                End If

                                int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)

                                'now enter Mean/Bias/n
                                If Count3 = 0 Then
                                    int8 = 0
                                    'Type Text labels for Statistics
                                    If boolAdHocStabCompColumns Then
                                        If Count2 = 0 Then
                                            Call typeStatsLabels(wd, int8, int1 - 1, 1, False)
                                        End If
                                    Else
                                        Call typeStatsLabels(wd, int8, int1 - 1, 1, False)
                                    End If


                                    If boolQCREPORTACCVALUES Then
                                    Else
                                        'intExp = 1 'must force in order to report stability
                                        'NO! This forces the "Summary Statistics Excluding Outlier Values" heading!
                                        If intExp = 0 Then
                                        Else

                                            If Count2 = 0 Then
                                                boolFirst = True
                                            End If

                                            int8 = int8 + 1
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                            'enter some blank spaces to fool PageBreak function
                                            '.selection.typetext(Text:="  ")

                                            If boolOutHeadI Or boolAdHocStabCompColumns Then
                                            Else
                                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                                .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                                Try
                                                    .Selection.Cells.Merge()
                                                Catch ex As Exception

                                                End Try
                                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                                End With
                                                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                                boolOutHeadI = True
                                            End If

                                            'This isn't working
                                            'intExp = 1 'must force in order to report stability. Put it here instead

                                            'Type Text labels for Statistics
                                            '20180418 LEE:
                                            If boolAdHocStabCompColumns Then
                                            Else

                                                Call typeStatsLabels(wd, int8, int1, intCol, False)

                                                If Count3 = numSets - 1 And numSets <> 1 And boolColDiff Then
                                                    'do stability
                                                    int8 = int8 + 1
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                    int8 = int8 + 1

                                                    str1 = GetLegendTitle(intTableID, idTR) ' "%Difference"
                                                    'if there is an '=' sign, then remove it
                                                    str1 = Replace(str1, " = ", "", 1, -1, CompareMethod.Text)
                                                    str1 = Replace(str1, "= ", "", 1, -1, CompareMethod.Text)
                                                    str1 = Replace(str1, " =", "", 1, -1, CompareMethod.Text)

                                                    '.Selection.TypeText("Stability(%)")

                                                    .Selection.TypeText(str1)
                                                End If
                                            End If



                                        End If
                                    End If
                                End If

                                int8 = 0
                                If boolAdHocStabCompColumns Then
                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                Else
                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                End If

                                intDiffRow1 = int1 + int8

                                v1 = arrFP(1, int12)
                                v2 = arrFP(2, int12)


                                Try
                                    numMean = 0
                                    numMean1 = 0
                                    nE = rows2E.Length

                                    'If boolRCConc Then
                                    '    var1 = MeanDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                    'Else
                                    '    var1 = MeanDRArea(rows2E, "CONCENTRATION", False, "ALIQUOTFACTOR", False, boolX)
                                    'End If
                                    '20180720 LEE: MeanDR can accept Area
                                    If boolRCConc Then
                                        var1 = MeanDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                    ElseIf boolX Then
                                        var1 = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                    Else
                                        var1 = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                    End If

                                    If boolRCConc Then
                                        If boolLUseSigFigs Then
                                            numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                    Else
                                        If boolRCPARatio Then
                                            If boolLUseSigFigsAreaRatio Then
                                                var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                            Else
                                                var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                            End If
                                        Else
                                            If boolLUseSigFigsArea Then
                                                var2 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                            Else
                                                var2 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                            End If
                                        End If
                                        '*****
                                        numMean = var2

                                    End If

                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)

                                Catch ex As Exception

                                End Try

                                'record tblMeans

                                'Excluded
                                If boolRCConc Then
                                    var1 = MeanDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                ElseIf boolX Then
                                    var1 = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                Else
                                    var1 = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                End If

                                If boolRCConc Then
                                    If boolLUseSigFigs Then
                                        var2 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    Else
                                        var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                    End If
                                ElseIf boolRCPARatio Then
                                    If boolX = False Then
                                        If boolLUseSigFigsAreaRatio Then
                                            var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                        Else
                                            var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                        End If
                                    Else
                                        If boolLUseSigFigsArea Then
                                            var2 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var2 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                        End If
                                    End If
                                Else
                                    If boolLUseSigFigsArea Then
                                        var2 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var2 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If

                                End If

                                Dim nr As DataRow = tblMeans.NewRow
                                nr("boolAll") = 0
                                nr("Mean") = var2
                                nr("Run") = Count2 'zero-based
                                nr("Level") = Count3 'zero-based
                                tblMeans.Rows.Add(nr)

                                'now do all
                                If boolRCConc Then
                                    var1 = MeanDR(rows2, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                ElseIf boolX Then
                                    var1 = MeanDR(rows2, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                Else
                                    var1 = MeanDR(rows2, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                End If

                                If boolRCConc Then
                                    If boolLUseSigFigs Then
                                        var2 = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    Else
                                        var2 = RoundToDecimalRAFZ(var1, LSigFig)
                                    End If
                                ElseIf boolRCPARatio Then
                                    If boolX = False Then
                                        If boolLUseSigFigsAreaRatio Then
                                            var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                        Else
                                            var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                        End If
                                    Else
                                        If boolLUseSigFigsArea Then
                                            var2 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var2 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                        End If
                                    End If
                                Else
                                    If boolLUseSigFigsArea Then
                                        var2 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var2 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                    End If

                                End If

                                Dim nr1 As DataRow = tblMeans.NewRow
                                nr1("boolAll") = -1
                                nr1("Mean") = var2
                                nr1("Run") = Count2 'zero-based
                                nr1("Level") = Count3 'zero-based
                                tblMeans.Rows.Add(nr1)


                                If rows2.Length = rows2E.Length Then
                                    boolHasOut = False
                                Else
                                    boolHasOut = True
                                End If



                                If boolSTATSMEAN Then
                                    numMean = 0
                                    numMean1 = 0
                                    nE = rows2E.Length
                                    Try
                                        'enter Mean
                                        int8 = int8 + 1

                                        'If boolRCConc Then
                                        '    var1 = MeanDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                        'Else
                                        '    var1 = MeanDRArea(rows2E, "CONCENTRATION", False, "ALIQUOTFACTOR", False, boolX)
                                        'End If
                                        '20180720 LEE: MeanDR can accept Area
                                        If boolRCConc Then
                                            var1 = MeanDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                        ElseIf boolX Then
                                            var1 = MeanDR(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                        Else
                                            var1 = MeanDR(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                        End If

                                        If boolRCConc Then
                                            If boolLUseSigFigs Then
                                                numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                            Else
                                                numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                            End If

                                            numMean1 = var1

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
                                                    .Selection.TypeText(CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                boolEnterDiff = True
                                            Else
                                                '.Selection.TypeText(Text:=CStr(numMean))
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If
                                                boolEnterDiff = True
                                            End If

                                        Else

                                            '*****
                                            If boolRCPARatio Then
                                                If boolLUseSigFigsAreaRatio Then
                                                    var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                Else
                                                    var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                                End If
                                            Else
                                                If boolLUseSigFigsArea Then
                                                    var2 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                Else
                                                    var2 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                                End If
                                            End If
                                            '*****
                                            numMean = var2
                                            numMean1 = var1
                                            .Selection.TypeText(CStr(var2))
                                        End If
                                        'numMean = SigFigOrDecString(RoundToDecimal(var1, 5), LSigFig, False)
                                        'if count2=0 then
                                        'If Count10 = 0 Then

                                        '1=level1 normal stats, 2 = level2 normal stats, 3=level1 2nd stats, 4 = level2 2nd stats, 
                                        '5=level1 normal stats fullprec, 6 = level2 normal stats fullprec, 7=level1 2nd stats fullprec, 8 = level2 2nd stats fullprec
                                        If boolAdHocStabCompColumns Then
                                            If Count2 = 0 Then
                                                arrS(1, Count2) = numMean
                                                arrS(5, Count2) = numMean1
                                                numT01 = numMean
                                            Else
                                                arrS(1, Count2) = numT01
                                                arrS(2, Count2) = numMean
                                                arrS(6, Count2) = numMean1
                                            End If
                                        Else
                                            'If Count2 = 0 Then
                                            '    arrS(1, int12 + 1) = numMean
                                            '    arrS(5, int12 + 1) = numMean1
                                            '    numT01 = numMean
                                            'Else
                                            '    'arrS(2, int12 + 1) = numMean
                                            '    arrS(2, int12 + 1) = numT01
                                            '    arrS(6, int12 + 1) = numMean1
                                            'End If

                                            If Count2 = 0 Then
                                                arrS(1, Count2) = numMean
                                                arrS(5, Count2) = numMean1
                                                numT01 = numMean
                                            Else
                                                'arrS(2, int12 + 1) = numMean
                                                arrS(1, Count2) = numT01
                                                arrS(2, Count2) = numMean
                                                arrS(6, Count2) = numMean1
                                            End If
                                        End If

                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If



                                Try
                                    If boolRCConc Then
                                        var1 = StdDevDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                    ElseIf boolX Then
                                        var1 = StdDevDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                    Else
                                        var1 = StdDevDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                    End If

                                    numSD = 0
                                    If boolRCConc Then
                                        If boolLUseSigFigs Then
                                            numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                        Else
                                            numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                        End If

                                    Else
                                        '*****
                                        If boolRCPARatio Then
                                            If boolLUseSigFigsAreaRatio Then
                                                var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                            Else
                                                var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                            End If
                                        Else
                                            If boolLUseSigFigsArea Then
                                                var1 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                            Else
                                                var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                            End If
                                            var2 = var1
                                        End If
                                        '*****
                                        numSD = var2 'RoundToDecimal(var1, 5)

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

                                            If boolRCConc Then
                                                var1 = StdDevDR(rows2E, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                            ElseIf boolX Then
                                                var1 = StdDevDRArea(rows2E, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                            Else
                                                var1 = StdDevDRArea(rows2E, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                            End If
                                            numSD = 0
                                            If boolRCConc Then
                                                If boolLUseSigFigs Then
                                                    numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                    .Selection.TypeText(CStr(DisplayNum(numSD, LSigFig, False)))
                                                Else
                                                    numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                                    .Selection.TypeText(CStr(Format(numSD, GetRegrDecStr(LSigFig))))
                                                End If

                                            Else
                                                '*****
                                                If boolRCPARatio Then
                                                    If boolLUseSigFigsAreaRatio Then
                                                        var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                    Else
                                                        var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                                    End If
                                                Else
                                                    If boolLUseSigFigsArea Then
                                                        var1 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                    Else
                                                        var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                                    End If
                                                    var2 = var1
                                                End If
                                                '*****
                                                numSD = var2 'RoundToDecimal(var1, 5)
                                                .Selection.TypeText(CStr(var2))
                                            End If

                                        End If

                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
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
                                        If nE < gSDMax Then
                                            .Selection.TypeText("NA")
                                        Else
                                            .Selection.TypeText(Format(numPrec, strQCDec))
                                        End If

                                        '.Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + 2).Select()
                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
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

                                        If boolRCPA Or boolRCPARatio Then
                                            .Selection.TypeText("NA")
                                        Else
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If
                                        '.Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + 2).Select()
                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
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
                                        If boolRCPA Or boolRCPARatio Then
                                            .Selection.TypeText("NA")
                                        Else
                                            .Selection.TypeText(Format(numTheor, strQCDec))
                                        End If
                                        '.Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + 2).Select()
                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
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

                                        If boolRCPA Or boolRCPARatio Then
                                            .Selection.TypeText("NA")
                                        Else
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If
                                        '.Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + 2).Select()
                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
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

                                        If boolRCPA Or boolRCPARatio Then
                                            .Selection.TypeText("NA")
                                        Else
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        End If

                                        If boolAdHocStabCompColumns Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
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
                                        .Selection.TypeText(CStr(nE))

                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolQCREPORTACCVALUES Then
                                Else
                                    nI = rows2A.Length
                                    'NO!
                                    'intExp = 1 'must force in order to report stability

                                    '20180419 LEE: Remove this. Do it at the end
                                    'If intExp = 0 Or boolAdHocStabCompColumns Then
                                    '    'If boolDeleteRows Then
                                    '    'Else
                                    '    '    Call DeleteRows(ctExp / intNumRuns, wd)
                                    '    '    boolDeleteRows = True
                                    '    'End If
                                    '    Call DeleteRows(ctExp, wd)
                                    'End If

                                    If intExp = 0 Then
                                        'If boolAdHocStabCompColumns Then
                                        '    int8 = 0
                                        'Else
                                        '    int8 = int8 + 2
                                        'End If
                                        'intDiffRow2 = int1 + int8
                                        intDiffRow2 = int1
                                    Else

                                        If boolAdHocStabCompColumns Then
                                            int8 = 0
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        Else
                                            int8 = int8 + 2
                                            .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                            'intDiffRow2 = int1 + int8
                                        End If
                                        intDiffRow2 = int1 + int8

                                        If boolSTATSMEAN Then
                                            Try
                                                'enter Mean
                                                int8 = int8 + 1

                                                'If boolRCConc Then
                                                '    var1 = MeanDR(rows2A, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                                'Else
                                                '    var1 = MeanDRArea(rows2A, "CONCENTRATION", False, "ALIQUOTFACTOR", False, boolX)
                                                'End If
                                                '20180720 LEE: MeanDR can accept Area
                                                If boolRCConc Then
                                                    var1 = MeanDR(rows2A, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                                ElseIf boolX Then
                                                    var1 = MeanDR(rows2A, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                                Else
                                                    var1 = MeanDR(rows2A, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                                End If
                                                If boolRCConc Then
                                                    If boolLUseSigFigs Then
                                                        numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                    Else
                                                        numMean = RoundToDecimalRAFZ(var1, LSigFig)
                                                    End If

                                                    numMean1 = var1

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
                                                        If boolAdHocStabCompColumns Then
                                                            str1 = .Selection.Text
                                                            If boolLUseSigFigs Then
                                                                str2 = CStr(DisplayNum(numMean, LSigFig, False))
                                                            Else
                                                                str2 = CStr(Format(numMean, GetRegrDecStr(LSigFig)))
                                                            End If
                                                            .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))
                                                            Else
                                                                .Selection.TypeText(CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                            End If

                                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                        End If

                                                        boolEnterDiff = True
                                                    Else
                                                        '.Selection.TypeText(Text:=CStr(numMean))
                                                        If boolAdHocStabCompColumns Then
                                                            str1 = .Selection.Text
                                                            If boolLUseSigFigs Then
                                                                str2 = CStr(DisplayNum(numMean, LSigFig, False))
                                                            Else
                                                                str2 = CStr(Format(numMean, GetRegrDecStr(LSigFig)))
                                                            End If
                                                            .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))
                                                            Else
                                                                .Selection.TypeText(CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                            End If
                                                        End If


                                                        boolEnterDiff = True
                                                    End If

                                                Else

                                                    '*****
                                                    If boolRCPARatio Then
                                                        If boolLUseSigFigsAreaRatio Then
                                                            var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                        Else
                                                            var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDec)
                                                        End If
                                                    Else
                                                        If boolLUseSigFigsArea Then
                                                            var2 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                        Else
                                                            var2 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                                        End If
                                                    End If
                                                    '*****
                                                    numMean = var2
                                                    numMean1 = var1
                                                    If boolAdHocStabCompColumns Then
                                                        str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                        str2 = CStr(var2)
                                                        .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                    Else
                                                        .Selection.TypeText(CStr(var2))
                                                    End If

                                                End If
                                                'numMean = SigFigOrDecString(RoundToDecimal(var1, 5), LSigFig, False)
                                                'if Count2 = 0
                                                '1=level1 normal stats, 2 = level2 normal stats, 3=level1 2nd stats, 4 = level2 2nd stats, 
                                                '5=level1 normal stats fullprec, 6 = level2 normal stats fullprec, 7=level1 2nd stats fullprec, 8 = level2 2nd stats fullprec
                                                If Count5 = 0 Then
                                                    arrS(3, Count2) = numMean
                                                    arrS(7, Count2) = numMean1
                                                Else
                                                    arrS(4, Count2) = numMean
                                                    arrS(8, Count2) = numMean1
                                                End If

                                                If boolAdHocStabCompColumns Then
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                                End If

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

                                                    If boolRCConc Then
                                                        var1 = StdDevDR(rows2A, "CONCENTRATION", True, "ALIQUOTFACTOR", True, boolX)
                                                    ElseIf boolX Then
                                                        var1 = StdDevDRArea(rows2A, "INTERNALSTANDARDAREA", False, "ALIQUOTFACTOR", False, boolX)
                                                    Else
                                                        var1 = StdDevDRArea(rows2A, "ANALYTEAREA", False, "ALIQUOTFACTOR", False, boolX)
                                                    End If
                                                    If boolRCConc Then

                                                        If boolAdHocStabCompColumns Then
                                                            str1 = .Selection.Text
                                                            If boolLUseSigFigs Then
                                                                numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                                str2 = CStr(DisplayNum(numSD, LSigFig, False))
                                                            Else
                                                                numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                                                str2 = CStr(Format(numSD, GetRegrDecStr(LSigFig)))
                                                            End If
                                                            .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                        Else
                                                            If boolLUseSigFigs Then
                                                                numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                                                .Selection.TypeText(CStr(DisplayNum(numSD, LSigFig, False)))
                                                            Else
                                                                numSD = RoundToDecimalRAFZ(var1, LSigFig)
                                                                .Selection.TypeText(CStr(Format(numSD, GetRegrDecStr(LSigFig))))
                                                            End If
                                                        End If


                                                    Else

                                                        '*****
                                                        If boolRCPARatio Then
                                                            If boolLUseSigFigsAreaRatio Then
                                                                var2 = SigFigAreaRatio(RoundToDecimalA(var1, LSigFigAreaRatio), LSigFigAreaRatio, False, True)
                                                            Else
                                                                var2 = Format(RoundToDecimalRAFZ(var1, LSigFigAreaRatio), strAreaDecAreaRatio)
                                                            End If
                                                        Else
                                                            If boolLUseSigFigsArea Then
                                                                var1 = SigFigArea(RoundToDecimalA(var1, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                                            Else
                                                                var1 = Format(RoundToDecimalRAFZ(var1, LSigFigArea), strAreaDec)
                                                            End If
                                                            var2 = var1
                                                        End If
                                                        '*****

                                                        numSD = var2

                                                        If boolAdHocStabCompColumns Then
                                                            str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                            str2 = CStr(var2)
                                                            .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                        Else
                                                            .Selection.TypeText(CStr(var2))
                                                        End If


                                                    End If

                                                End If

                                                If boolAdHocStabCompColumns Then
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If boolSTATSCV Then
                                            Try
                                                'enter %CV
                                                int8 = int8 + 1
                                                If nI < gSDMax Then
                                                    If boolAdHocStabCompColumns Then
                                                        str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                        str2 = "NA"
                                                        .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                    Else
                                                        .Selection.TypeText("NA")
                                                    End If
                                                Else
                                                    numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                                    If boolAdHocStabCompColumns Then
                                                        str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                        str2 = Format(numPrec, strQCDec)
                                                        .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                    Else
                                                        .Selection.TypeText(Format(numPrec, strQCDec))
                                                    End If
                                                End If

                                                If boolAdHocStabCompColumns Then
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If
                                        If boolSTATSBIAS And boolSTATSMEAN Then
                                            Try
                                                'enter %Bias
                                                int8 = int8 + 1
                                                numBias = CalcREPercent(numMean, varNom, intQCDec)

                                                If boolAdHocStabCompColumns Then
                                                    str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                    str2 = Format(numBias, strQCDec)
                                                    .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                Else
                                                    If boolRCPA Or boolRCPARatio Then
                                                        .Selection.TypeText("NA")
                                                    Else
                                                        .Selection.TypeText(Format(numBias, strQCDec))
                                                    End If
                                                End If


                                                If boolAdHocStabCompColumns Then
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
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
                                                '.Selection.TypeText(CStr(var1))

                                                numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                                numTheor = 100 + CDec(numTheor)

                                                If boolAdHocStabCompColumns Then

                                                    str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                    str2 = Format(numTheor, strQCDec)
                                                    .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")

                                                Else
                                                    If boolRCPA Or boolRCPARatio Then
                                                        .Selection.TypeText("NA")
                                                    Else
                                                        .Selection.TypeText(Format(numTheor, strQCDec))
                                                    End If
                                                End If

                                                If boolAdHocStabCompColumns Then
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                                End If
                                            Catch ex As Exception

                                            End Try

                                        End If

                                        If boolSTATSDIFF And boolSTATSMEAN Then
                                            Try
                                                'enter %Bias
                                                int8 = int8 + 1
                                                numBias = CalcREPercent(numMean, varNom, intQCDec)

                                                If boolAdHocStabCompColumns Then

                                                    str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                    str2 = Format(numBias, strQCDec)
                                                    .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")

                                                Else
                                                    If boolRCPA Or boolRCPARatio Then
                                                        .Selection.TypeText("NA")
                                                    Else
                                                        .Selection.TypeText(Format(numBias, strQCDec))
                                                    End If
                                                End If

                                                If boolAdHocStabCompColumns Then
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If BOOLSTATSRE And boolSTATSMEAN Then
                                            Try
                                                'enter %RE
                                                int8 = int8 + 1
                                                numBias = CalcREPercent(numMean, varNom, intQCDec)

                                                If boolAdHocStabCompColumns Then

                                                    str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                    str2 = Format(numBias, strQCDec)
                                                    .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")

                                                Else
                                                    If boolRCPA Or boolRCPARatio Then
                                                        .Selection.TypeText("NA")
                                                    Else
                                                        .Selection.TypeText(Format(numBias, strQCDec))
                                                    End If
                                                End If

                                                If boolAdHocStabCompColumns Then
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, (Count3 * intDiffInc) + intColSN + 1).Select()
                                                End If
                                            Catch ex As Exception

                                            End Try
                                        End If

                                        If boolSTATSN Then
                                            Try
                                                'enter n
                                                int8 = int8 + 1

                                                If boolAdHocStabCompColumns Then

                                                    str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                    str2 = CStr(nI)
                                                    .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")

                                                Else
                                                    .Selection.TypeText(CStr(nI))
                                                End If



                                            Catch ex As Exception

                                            End Try
                                        End If

                                    End If
                                    'End If

                                End If

                                var1 = var1

                            Next Count3 'start filling in data by columns, intRowsX = 0, For Count3 = 0 To (intNumLevels * int11) - 1 Step int11 

                            intoStart = int1 + int8 + 2
                            intStart = intoStart

                            'int8 = int8 + 1
                            If boolAdHocStabCompColumns And boolColDiff Then
                                int8 = int8 + 1
                                If Count2 = 0 Then
                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol1).Select()
                                    .Selection.TypeText(strFirst)
                                End If
                            Else
                                'If Count2 = 0 Then
                                'Else
                                '    .Selection.Tables.Item(1).Cell(int1 + int8, 1).Select()
                                '    .Selection.TypeText(strFirst)
                                'End If
                            End If

                            If boolAdHocStabCompColumns Then
                                .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                            Else
                                .Selection.Tables.Item(1).Cell(intDiffRow1, intCol).Select()
                            End If

                            '20180724 LEE:
                            'New logic. No need to calculate %Diff if legend not shown
                            '20180727 LEE:
                            'Aaack! Bad logic! Grid-style requires %Diff
                            '20180807 LEE:
                            ''Still bad logic. Diff calcs need to be calculated no matter what
                            'For Count4 = 1 To intNumLevels
                            '    If Count2 = 0 Then
                            '    Else
                            '        '1=level1 normal stats, 2 = level2 normal stats, 3=level1 2nd stats, 4 = level2 2nd stats, 
                            '        '5=level1 normal stats fullprec, 6 = level2 normal stats fullprec, 7=level1 2nd stats fullprec, 8 = level2 2nd stats fullprec
                            '        If gboolMeanFullPrec Then
                            '            '20161206 LEE: gboolMeanFullPrec has been deprecated. Only gboolMeanRounded = true (e.g. gboolMeanFullPrec = false) is allowed
                            '            var1 = arrS(5, Count2) 'old
                            '            var2 = NZ(arrS(6, Count2), 1) 'new
                            '        Else
                            '            var1 = numT01 'arrS(1, Count2) 'old
                            '            var2 = NZ(arrS(2, Count2), 1) 'new
                            '        End If
                            '        var3 = ReturnDiff(CDec(var1), CDec(var2))
                            '        Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count4, "Diff", var3, CSng(var10), Count1, strDo, 0, 0, False)
                            '    End If
                            'Next Count4

                            '20180807 LEE
                            intRunM = intRunM + 1

                            'If boolNONELEG And boolAdHocStabCompColumns = False Then
                            '20181220 LEE
                            'do not evaluate boolNONELEG here anymore
                            If boolAdHocStabCompColumns = True Then
                                .Selection.Tables.Item(1).Cell(int1 + int8 - 1, intCol).Select()
                            End If

                            For Count4 = 1 To intNumLevels
                                'If Count2 = intNumRuns - 1 And intNumRuns > 1 Then
                                'enter stability
                                '1=level1 normal stats, 2 = level2 normal stats, 3=level1 2nd stats, 4 = level2 2nd stats,

                                var10 = 1
                                varNom = tblLevels.Rows.Item(Count4 - 1).Item("NOMCONC")

                                '1=level1 normal stats, 2 = level2 normal stats, 3=level1 2nd stats, 4 = level2 2nd stats, 
                                '5=level1 normal stats fullprec, 6 = level2 normal stats fullprec, 7=level1 2nd stats fullprec, 8 = level2 2nd stats fullprec
                                If gboolMeanFullPrec Then
                                    '20161206 LEE: gboolMeanFullPrec has been deprecated. Only gboolMeanRounded = true (e.g. gboolMeanFullPrec = false) is allowed
                                    var1 = arrS(5, Count2) 'old
                                    var2 = NZ(arrS(6, Count2), 1) 'new
                                Else
                                    var1 = numT01 'arrS(1, Count2) 'old
                                    var2 = NZ(arrS(2, Count2), 1) 'new
                                End If

                                '20180807 LEE:
                                'Retrieve Means
                                Dim strFM As String
                                'get t0
                                strFM = "boolAll = 0 and Run = 0 AND LEVEL = " & Count4 - 1
                                rowsMean = tblMeans.Select(strFM)
                                If rowsMean.Length = 0 Then
                                    numMeanMT0 = 0
                                Else
                                    numMeanMT0 = rowsMean(0).Item("Mean")
                                End If
                                var1 = numMeanMT0

                                '20190117 LEE:
                                'need rowsAll for T0
                                strFM = "boolAll = -1 and Run = 0 AND LEVEL = " & Count4 - 1
                                rowsMeanT0All = tblMeans.Select(strFM)
                                If rowsMeanT0All.Length = 0 Then
                                    numMeanT0All = 0
                                Else
                                    numMeanT0All = rowsMeanT0All(0).Item("Mean")
                                End If


                                'get current value
                                strFM = "boolAll = 0 and Run = " & intRunM & " AND LEVEL = " & Count4 - 1
                                rowsMean1 = tblMeans.Select(strFM)
                                If rowsMean1.Length = 0 Then
                                    numMeanM = 0
                                Else
                                    numMeanM = rowsMean1(0).Item("Mean")
                                End If
                                var2 = numMeanM

                                If IsDBNull(var1) Or IsDBNull(var2) Then
                                    var3 = "NA"
                                Else
                                    If Len(var1) = 0 Or Len(var2) = 0 Then
                                        var3 = "NA"
                                    Else
                                        If var1 = 0 Then
                                            var3 = "NA"
                                        Else
                                            'Here
                                            var3 = 0

                                            var3 = ReturnDiff(CDec(var1), CDec(var2))

                                            If Count2 = 0 Then
                                            Else
                                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count4, "Diff", var3, CSng(var10), Count1, strDo, 0, 0, False)
                                            End If

                                        End If

                                    End If

                                    If boolAdHocStabCompColumns Then
                                        If boolColDiff Then
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                        End If

                                    Else

                                        If boolSTATSDIFFCOL Then
                                            .Selection.Tables.Item(1).Cell(intDiffRow1, (intNumLevels * 2) + intColSN + 1).Select()
                                        Else


                                            If tbl.Columns.Count >= intNumLevels + intColSN + 1 Then
                                                '20180807 LEE:
                                                If boolColDiff Then
                                                    .Selection.Tables.Item(1).Cell(intDiffRow1, intColSN + (Count4 * 2)).Select()
                                                Else
                                                    .Selection.Tables.Item(1).Cell(intDiffRow1, intNumLevels + intColSN + 1).Select()
                                                End If

                                            Else
                                                .Selection.Tables.Item(1).Cell(intDiffRow1, tbl.Columns.Count).Select()
                                            End If

                                        End If

                                    End If

                                    '20181220 LEE
                                    If boolColDiff Then
                                        If boolAdHocStabCompColumns Then
                                            If Count2 = 0 Then
                                                .Selection.TypeText("NA")
                                                boolNA = True
                                            Else
                                                .Selection.TypeText(CStr(Format(var3, strQCDec)))
                                            End If

                                        Else
                                            If Count2 = 0 Then
                                                .Selection.TypeText("NA")
                                                boolNA = True
                                                int8 = int8 - 1

                                            Else
                                                '20190117 LEE:
                                                If numMeanMT0 = numMeanT0All Then
                                                    .Selection.TypeText(CStr(Format(var3, strQCDec)))
                                                Else
                                                    'add additional diff
                                                    var4 = CStr(Format(ReturnDiff(CDec(numMeanT0All), CDec(var2)), strQCDec))
                                                    var3 = CStr(Format(var3, strQCDec)) & "(" & var4 & ")"
                                                    .Selection.TypeText(var3)
                                                End If


                                            End If

                                            'need to go to last stats row for end of data bordering
                                            .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()

                                        End If

                                    End If

                                End If

                                If boolQCREPORTACCVALUES Then
                                Else

                                    If Count2 = 0 Then

                                        If intExp > 0 Then

                                            If boolSTATSDIFFCOL Then
                                                .Selection.Tables.Item(1).Cell(intDiffRow2, (intNumLevels * 2) + intColSN + 1).Select()
                                            Else


                                                If tbl.Columns.Count >= intNumLevels + intColSN + 1 Then
                                                    '20180807 LEE:
                                                    If boolColDiff Then
                                                        .Selection.Tables.Item(1).Cell(intDiffRow2, intColSN + (Count4 * 2)).Select()
                                                    Else
                                                        .Selection.Tables.Item(1).Cell(intDiffRow2, intNumLevels + intColSN + 1).Select()
                                                    End If

                                                Else
                                                    .Selection.Tables.Item(1).Cell(intDiffRow2, tbl.Columns.Count).Select()
                                                End If

                                            End If

                                            .Selection.TypeText("NA")
                                            boolNA = True
                                        End If
                                    Else

                                        If intExp > 0 Then

                                            '1=level1 normal stats, 2 = level2 normal stats, 3=level1 2nd stats, 4 = level2 2nd stats, 
                                            '5=level1 normal stats fullprec, 6 = level2 normal stats fullprec, 7=level1 2nd stats fullprec, 8 = level2 2nd stats fullprec
                                            If gboolMeanFullPrec Then
                                                '20161206 LEE: gboolMeanFullPrec has been deprecated. Only gboolMeanRounded = true (e.g. gboolMeanFullPrec = false) is allowed
                                                var1 = arrS(5, Count2) 'old
                                                var2 = NZ(arrS(6, Count2), 1) 'new
                                            Else
                                                var1 = numT01 'arrS(1, Count2) 'old
                                                var2 = NZ(arrS(4, Count2), 1) 'new
                                            End If


                                            '20180807 LEE:
                                            'Retrieve Means
                                            Dim strFMa As String
                                            'get t0
                                            strFMa = "boolAll = -1 and Run = 0 AND LEVEL = " & Count4 - 1
                                            rowsMean = tblMeans.Select(strFMa)
                                            If rowsMean.Length = 0 Then
                                                numMeanMAllT0 = 0
                                            Else
                                                numMeanMAllT0 = rowsMean(0).Item("Mean")
                                            End If
                                            var1 = numMeanMAllT0

                                            'get current value
                                            strFMa = "boolAll = -1 and Run = " & intRunM & " AND LEVEL = " & Count4 - 1
                                            rowsMeanAll = tblMeans.Select(strFMa)
                                            If rowsMeanAll.Length = 0 Then
                                                numMeanM = 0
                                            Else
                                                numMeanM = rowsMeanAll(0).Item("Mean")
                                            End If
                                            var2 = numMeanM

                                            If IsDBNull(var1) Or IsDBNull(var2) Then
                                                var3 = "NA"
                                            Else
                                                If Len(var1) = 0 Or Len(var2) = 0 Then
                                                    var3 = "NA"
                                                Else
                                                    If var1 = 0 Then
                                                        var3 = "NA"
                                                    Else
                                                        'Here
                                                        var3 = 0

                                                        var3 = ReturnDiff(CDec(var1), CDec(var2))

                                                        'If Count2 = 0 Then
                                                        'Else
                                                        '    'Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count4, "Diff", var3, CSng(var10), Count1, strDo, 0, 0, False)
                                                        'End If

                                                    End If

                                                End If

                                                If boolAdHocStabCompColumns Then
                                                    If boolColDiff Then
                                                        .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()
                                                    End If

                                                Else

                                                    'If boolSTATSDIFFCOL Then
                                                    '    .Selection.Tables.Item(1).Cell(intDiffRow2, (intNumLevels * 2) + intColSN + 1).Select()
                                                    'Else
                                                    '    .Selection.Tables.Item(1).Cell(intDiffRow2, intNumLevels + intColSN + 1).Select()
                                                    'End If

                                                    '20190225 LEE:
                                                    If boolSTATSDIFFCOL Then
                                                        .Selection.Tables.Item(1).Cell(intDiffRow2, (intNumLevels * 2) + intColSN + 1).Select()
                                                    Else


                                                        If tbl.Columns.Count >= intNumLevels + intColSN + 1 Then
                                                            '20180807 LEE:
                                                            If boolColDiff Then
                                                                .Selection.Tables.Item(1).Cell(intDiffRow2, intColSN + (Count4 * 2)).Select()
                                                            Else
                                                                .Selection.Tables.Item(1).Cell(intDiffRow2, intNumLevels + intColSN + 1).Select()
                                                            End If

                                                        Else
                                                            .Selection.Tables.Item(1).Cell(intDiffRow2, tbl.Columns.Count).Select()
                                                        End If

                                                    End If

                                                End If

                                                str1 = Mid(.Selection.Text, 1, Len(.Selection.Text) - 2)
                                                If boolAdHocStabCompColumns Then
                                                    If boolColDiff Then
                                                        If Count2 = 0 Then
                                                            str2 = "NA"
                                                            .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                            boolNA = True
                                                        Else
                                                            str2 = CStr(Format(var3, strQCDec))
                                                            .Selection.TypeText(str1 & ChrW(160) & "(" & str2 & ")")
                                                            '.Selection.TypeText(CStr(Format(var3, strQCDec)))
                                                        End If
                                                    End If

                                                Else
                                                    If Count2 = 0 Then
                                                        .Selection.TypeText("NA")
                                                        int8 = int8 - 1
                                                        boolNA = True
                                                    Else
                                                        .Selection.TypeText(CStr(Format(var3, strQCDec)))
                                                    End If

                                                    'need to go to last stats row for end of data bordering
                                                    .Selection.Tables.Item(1).Cell(int1 + int8, intCol).Select()

                                                End If

                                            End If

                                        End If

                                    End If

                                End If

                            Next Count4

                            If boolAdHocStabCompColumns = True Then
                                .Selection.Tables.Item(1).Cell(int1 + int8 - 1, intCol).Select()
                            Else

                            End If

                            var1 = var1 'debug

                        Next Count2 'this is H2 set

                        'If there is only one Run Identifier, then do not do Legend

                        'autofit window
                        .Selection.Tables.Item(1).Select()
                        Call autofitContent(wd, 2)

                        'go back and merge line 1
                        .Selection.Tables.Item(1).Cell(1, 2).Select()

                        If boolAdHocStabCompColumns Then

                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try


                            '.Selection.Tables.Item(1).Cell(2, 2).Select()
                            '20181127 LEE:
                            .Selection.Tables.Item(1).Cell(intLRow, 2).Select()

                        End If

                        '20180313 LEE:
                        'merge values and %Diff because if numLevels > 1, then str1 below must be extended to end of table row
                        'merge even if numLevels = 1 to be consistent
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.Cells.Merge()
                        Try
                            .Selection.Cells.Merge()
                        Catch ex As Exception

                        End Try


                        .Selection.Font.Bold = False
                        If boolRCConc Then
                            str1 = "Nominal Concentration"
                        ElseIf boolRCPA Then
                            str1 = "Peak Area"
                        ElseIf boolRCPARatio Then
                            If boolX Then 'is Int Std
                                str1 = "Peak Area"
                            Else
                                str1 = "Peak Area Ratio"
                            End If
                        Else
                            str1 = "Nominal Concentration"
                        End If
                        .Selection.TypeText(Text:=str1)

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
                    Call RemoveRows(wd, ctTbl)

                    str1 = str2 & " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."

                    '***
                    var1 = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    'strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    If boolX Then
                        strA = "Internal Standard " & rows11(Count1 - 1).Item("AnalyteDescription")
                    Else
                        strA = rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                    End If
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, boolX)

                    Call EnterTableNumber(wd, strTName, intCRow, strA, strTempInfo, intTableID, intGroup, idTR)

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

                    str1 = strM1

                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, ctLegend + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    'move to line below table
                    Call MoveOneCellDown(wd)

                    '20190108 LEE:
                    'boolNoneLeg evaluated in InsertLegend
                    'If boolNONELEG Or numSets = 1 Then
                    'Else
                    '    Call InsertLegend(wd, intTableID, idTR, boolX, 1)
                    'End If
                    If numSets = 1 Then
                    Else
                        Call InsertLegend(wd, intTableID, idTR, boolX, 1)
                    End If

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

                    str1 = str2 & " Final Extract Stability: Summary of " & rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION") & " Interpolated QC Standard Concentrations."
                    str2 = str1
                    str1 = NZ(rows11(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION"), "")
                    'Call JustTable(wd, str1, str2, strDo, strTName, intTableID)
                    'Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "")
                    strA = str1
                    If Len(str1) = 0 Then

                    Else
                        strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                        Call JustTable(wd, str1, strTName, strDo, strTName, intTableID, strTempInfo, "", strTNameO, intGroup, idTR)
                    End If

                End If

next1:

            Next Count1
end2:
        End With

    End Sub


    'check if table is to be generated (any of the Analyte Run Groups are selected FOR THIS MATRIX)
    Function boolGenerateTableForThisAnalyteIDandMatrix(ByVal intDo As Short, ByVal strAnalyteID As String,
                                                         ByVal strMatrix As String) As Boolean
        Dim Count1 As Short
        Dim strDo As String
        Dim dvDo As System.Data.DataView
        dvDo = frmH.dgvReportTableConfiguration.DataSource

        boolGenerateTableForThisAnalyteIDandMatrix = False
        For Count1 = 0 To tblAnalyteGroups.Rows.Count - 1
            'If the correct AnalyteID, check if selected
            If (StrComp(tblAnalyteGroups.Rows(Count1).Item("ANALYTEID"), strAnalyteID) = 0) Then 'Right AnalyteID
                If (StrComp(tblAnalyteGroups.Rows(Count1).Item("MATRIX"), strMatrix) = 0) Then 'Right Matrix
                    strDo = tblAnalyteGroups.Rows(Count1).Item("ANALYTEDESCRIPTION_C") 'Sub-Analyte name (eg. CmpdXYZ_C1)
                    If dvDo.Item(intDo).Item(strDo) Then 'If analyte is checked *for this table*
                        boolGenerateTableForThisAnalyteIDandMatrix = True  'We should generate this table
                        Exit For
                    End If
                End If
            End If
        Next

    End Function

    Function makeRunMatrixAndSelectedAnalyteFilter(ByVal intDo As Short, ByVal strAnalyteID As String, ByVal strMatrix As String)
        'Takes all the Accepted Runs, and create a SQL filter for only the Runs that contain the AnalyteID and Matrix given.
        'Then Filters those runs based on which ones have conditions (i.e. Calibration Sets) of Analyte Groups
        'which were not selected (checkbox was unchecked) on the Table Configurations table.
        makeRunMatrixAndSelectedAnalyteFilter = makeRunFilter(intDo, strAnalyteID, strMatrix, True)
    End Function

    Function makeRunMatrixFilter(ByVal intDo As Short, ByVal strAnalyteID As String, ByVal strMatrix As String)
        'Takes all the Accepted Runs, and create a SQL filter for only the Runs that contain the AnalyteID and Matrix given.
        makeRunMatrixFilter = makeRunFilter(intDo, strAnalyteID, strMatrix, False)
    End Function


    Function makeRunFilter(ByVal intDo As Short, ByVal strAnalyteID As String, ByVal strMatrix As String, _
                           ByVal useRunSelectionsForDifferingCalibrationSets As Boolean)
        'Description: see makeRunMatrixFilter() and makeRunMatrixAndSelectedAnalyteFilter()
        'NDL 5-Feb-2016:  We have to find a list of rows that are valid.  These are rows that have
        'the right AnalyteID and the right matrix.  In addition, sometimes we filter out the rows of certain
        'calibration sets if the "_C" Analyte Group is not checked for this table in the Table configuration.

        'Select for all runs that AnalyteIDs equal to this one
        Dim dv2 As New DataView(tblCalStdGroupAssayIDsAcc)
        Dim tblRuns As New System.Data.DataTable
        Dim strF, strDo As String
        Dim Count1 As Short
        Dim dvDo As System.Data.DataView

        dvDo = frmH.dgvReportTableConfiguration.DataSource

        'FILTER1:Filter the runs with the right Analyte ID and Matrix 
        strF = "ANALYTEID = " & strAnalyteID & " AND MATRIX = '" & strMatrix & "'"


        'FILTER2:Filter out the runs which have Sub-Analytes that are not selected
        If (useRunSelectionsForDifferingCalibrationSets) Then
            Dim strAnalyteDescriptionC As String
            For Count1 = 0 To tblAnalyteGroups.Rows.Count - 1 'go through all subAnalytes
                strAnalyteDescriptionC = tblAnalyteGroups.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
                If (Not (dvDo.Item(intDo).Item(strAnalyteDescriptionC)) And _
                    (StrComp(strAnalyteID, tblAnalyteGroups.Rows(Count1).Item("ANALYTEID")) = 0)) Then
                    'Right AnalyteID, but checkbox unchecked
                    strF = strF & " AND ANALYTEDESCRIPTION_C <> '" & strAnalyteDescriptionC & "' "
                End If
            Next
        End If

        'FILTER3:Create a Runs table with only the relevant runs
        dv2.RowFilter = strF
        tblRuns = dv2.ToTable("tblRuns", True, "ANALYTEID", "RUNID", "MATRIX", "ANALYTEDESCRIPTION_C")

        'FILTER4:Start tblSampleDesign by filtering for AnalyeID
        strF = "(ANALYTEID = " & strAnalyteID & ") "

        'FILTER5:Make a text filter to filter only the rows in tblRuns from tblSampleDesign
        Dim ctRuns As Short
        Dim strRuns As String

        If (tblRuns.Rows.Count > 0) Then
            strF = strF & " AND ("
        End If

        For ctRuns = 0 To tblRuns.Rows.Count - 1
            Dim var10 As String
            var10 = tblRuns.Rows(ctRuns).Item("RUNID")
            strF = strF & "(RUNID = " & var10 & ") "
            If ctRuns < tblRuns.Rows.Count - 1 Then
                strF = strF & " OR "
            Else
                strF = strF & ") "
            End If
        Next
        makeRunFilter = strF
    End Function

End Module
