Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports System.IO

Module modStyle4_UseGroups


    Sub SRSummaryOfBCSC_UseGroups_3(ByVal wd As Word.Application, ByVal intRunNum As Int16, ByVal rows() As Object, ByVal ctAnal As Short, ByVal intTableID As Int16, ByVal tblBCS As DataTable, ByVal tblBCSC As DataTable, ByVal idTR As Int64)


        'intRunNum(0) = from normal BCSC table
        'intRunNum(1) = from special BCSC table requires rows() of RunID's on Analyte ID ctAnal
        'tblNC = tblBCStds filtered for chosen assayids

        Dim boolOC As Boolean = False 'bool if eliminated
        Dim numNomConc As Decimal
        Dim BACStudy As String
        Dim rs As New ADODB.Recordset
        Dim constr As String
        Dim dbPath As String
        Dim str1, str2, str3, str4 As String
        'Dim arrAnalytes(1 To 7, 1 To 50) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns
        Dim Count1, Count2, Count3, Count4, Count5, Count6, Count7 As Short
        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9
        Dim int1, int2, int3, int8 As Short
        Dim arrTemp(2, 50)
        Dim num1, num2 As Object
        Dim num3 As Double
        Dim arrBCStdActual(1)

        Dim ctLegend As Short
        Dim lng1, lng2 As Integer
        Dim boolPortrait As Boolean
        Dim intLastAnal As Short
        Dim arrOrder()
        Dim ctCols 'number of columns in a table
        Dim strSub1, strSub2 As String
        Dim pos1, pos2 As Short
        Dim numSum As Object
        Dim numMean As Object
        Dim numSD As Object
        Dim dvDo As System.Data.DataView
        Dim intDo As Short
        Dim strDo As String
        Dim bool As Boolean
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim dv As System.Data.DataView
        Dim drows() As DataRow
        Dim ctP As Short
        Dim arrBCStds(2, 20) '1=LevelNumber, 2=Concentration
        Dim arrBCStdConcs(10, 10)
        Dim inttemprows As Short
        Dim strTName As String
        Dim intS, intE, intN As Short
        Dim strS, strA, strF As String
        Dim strF1 As String
        Dim strTempInfo As String
        Dim intCStdReps As Short
        Dim rowsCSR() As DataRow
        Dim intNumSamples As Short
        Dim tblBCStds1, tblBCStdConcs1 As System.Data.DataTable
        Dim boolEx As Boolean
        Dim strDecReason As String

        Dim boolHit As Boolean
        Dim boolNoDiff As Boolean

        Dim intLeg, intExp, ctExp As Short

        Dim rowsAssID() As DataRow

        Dim fontsize
        Dim boolPro As Boolean
        Dim intLegStart As Short

        Dim strRegressionType As String
        Dim strWeighting As String
        Dim intRP As Short
        Dim intTRows As Short
        Dim arrRegCon(1, 1)

        Dim intRPIncr As Short

        Dim arrRegr(5, 10) '1=A, 2=B, n=R2
        Dim intRegrCt As Short

        Dim varNom
        Dim strConcUnits As String

        Dim arrRunID(1)
        Dim intRunID As Int64
        Dim boolSRegr As Boolean = True 'Single Regression
        Dim boolSWt As Boolean = True 'Single Weighting
        Dim intNumRegr As Short
        Dim arrRegrType(2, 1) 'need for legend 1=Type, 2=Wt

        Dim intNumRuns As Short
        Dim boolAssSamps As Boolean = False

        Dim intColCt As Short

        Dim strNRA As String = "NRA"
        Dim strNRB As String = "NRB"
        Dim strNR As String = "NR"
        Dim intNR As Short
        Dim arrNR(3, 200)
        Dim arrNRU(2)
        Dim intNRPos As Short = 0
        Dim intNRUsed As Short = 0
        Dim numAFP

        Dim strFFF As String

        Dim intLegNNR As Short = 0

        Dim hi As Double
        Dim lo As Double

        Dim v1, v2, vU

        Dim numPrec As Single
        Dim numBias As Single
        Dim numTheor As Single

        Dim intAssayID As Int64

        Dim boolRAID As Boolean = False

        Dim charFCID As String
        Dim boolNotAssignedSamples = False

        Dim intGroup As Short
        Dim intAnalyteIndex As Int64
        Dim intAnalyteID As Int64
        Dim strAnalayteFlag As String
        Dim strMatrix As String

        Dim intSpecies As Short
        Dim intCR As Short

        Dim strTNameO As String 'original Table Name

        Dim strFAssayID As String

        If (intRunNum = 0) Then
            boolNotAssignedSamples = True
        End If

        '* Choose the correct Report Table (passed into function)
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        '* Calculate and store tables of Back-Calculated Standards, #Analytes, Concentration Levels, etc.
        Dim strLevelNum As String
        Dim strNomConc As String
        If boolUseGroups Then
            If boolNotAssignedSamples Then
                intS = 1  'intS, intE  = Start & End analytes
                intE = ctAnalytes 'Total number of Analytes (global variable)
                intTableID = 3
                tblBCStds1 = tblCalStdGroupsAcc 'tblBCStds
                tblBCStdConcs1 = tblBCStdConcs
                boolEx = False
                strLevelNum = "LEVELNUMBER"
                strNomConc = "CONCENTRATION"
            Else 'Samples Assigned in Assign Samples... window
                intS = ctAnal  'Number of Analytes (passed into Function)
                intE = intS  'intS, intE  = Start & End analytes
                tblBCStds1 = tblBCS 'tblNC2
                tblBCStdConcs1 = tblBCSC 'tblNC3 = filtered tblBCStdConcs
                boolEx = True
                strLevelNum = "ASSAYLEVEL"
                strNomConc = "NOMCONC"
            End If
        Else
            If boolNotAssignedSamples Then
                intS = 1  'intS, intE  = Start & End analytes
                intE = ctAnalytes 'Total number of Analytes (global variable)
                intTableID = 3
                tblBCStds1 = tblBCStds
                tblBCStdConcs1 = tblBCStdConcs
                boolEx = False
                strLevelNum = "LEVELNUMBER"
                strNomConc = "CONCENTRATION"
            Else 'Samples Assigned in Assign Samples... window
                intS = ctAnal  'Number of Analytes (passed into Function)
                intE = intS  'intS, intE  = Start & End analytes
                tblBCStds1 = tblBCS 'tblNC2
                tblBCStdConcs1 = tblBCSC 'tblNC3 = filtered tblBCStdConcs
                boolEx = True
                strLevelNum = "ASSAYLEVEL"
                strNomConc = "NOMCONC"
            End If
        End If

        'in this routine, intTableID may come passed in the routine parameters
        Dim strWRunId As String = GetWatsonColH(intTableID)

        '* Make table of Nominal Concentration & AssayIDs 
        Dim tblRunIDNomConc As New System.Data.DataTable
        Dim col10 As New DataColumn

        tblRunIDNomConc.Columns.Add("NomConc", Type.GetType("System.Single"))
        tblRunIDNomConc.Columns.Add("AssayID", Type.GetType("System.Int64"))

        '* Grab Information from Advanced Table Configuration table
        dvDo = frmH.dgvReportTableConfiguration.DataSource
        intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")
        var1 = dvDo(intDo).Item("CHARHEADINGTEXT") 'Get table name

        strTNameO = NZ(var1, "[NONE]")

        var1 = dvDo(intDo).Item("CHARSTABILITYPERIOD") 'get Temperature info
        strTempInfo = NZ(var1, "[NONE]")

        '* Update Progress Display
        ctPB = ctPB + 1
        If ctPB > frmH.pb1.Maximum Then
            ctPB = 1
        End If
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()

        '''''''''''''''wdd.visible = True

        Dim fonts
        fontsize = wd.ActiveDocument.Styles("Normal").Font.Size 'wd.Selection.Font.Size
        fonts = fontsize ' wd.Selection.Font.Size

        Dim strM1 As String
        ''''''''wdd.visible = True

        '20150303 New Methodology
        'Must be:
        'AnalyteID
        '    AssayID
        '        AnalyticalRun: filtered by AssayID, AnalyteIndex, MasterAssayID

        'so find all this stuff here:

        '* For each Analyte, Write out a Table

        Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted
        Dim strAnal As String
        Dim strAnalC As String
        Dim rowsAssays() As DataRow
        Dim strFAID As String

        ''tblBCStdsAssayID
        ''debug
        'Console.WriteLine("Start tblBCStdsAssayID")
        'var1 = ""
        'For Count1 = 0 To tblBCStdsAssayID.Columns.Count - 1
        '    var2 = tblBCStdsAssayID.Columns(Count1).ColumnName
        '    var1 = var1 & ChrW(9) & var2
        'Next
        'Console.WriteLine(var1)
        'For Count2 = 0 To tblBCStdsAssayID.Rows.Count - 1
        '    var1 = ""
        '    For Count1 = 0 To tblBCStdsAssayID.Columns.Count - 1
        '        var2 = tblBCStdsAssayID.Rows(Count2).Item(Count1)
        '        var1 = var1 & ChrW(9) & var2
        '    Next
        '    Console.WriteLine(var1)
        'Next
        'Console.WriteLine("End tblBCStdsAssayID")

        With wd   'Word document shortcut
            For Count1 = intS To intE 'For all Analytes to have tables (intS, intE  = Start & End analytes)...

                strTName = strTNameO 'reset strTName

                ctLegend = 0
                Dim arrLegend(4, 500)

                tblRunIDNomConc.Clear()

                '* Calculate columns per nominal mass value (int11)
                Dim int11 As Short
                If boolSTATSDIFFCOL Then
                    int11 = 2
                Else
                    int11 = 1
                End If

                intLeg = 0
                intLegNNR = 0
                intExp = 0
                intLegStart = 96

                '* Check if table is to be generated for this Analyte
                strAnal = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION")
                strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")
                strDo = strAnalC ' arrAnalytes(1, Count1) 'record column name (Analyte Description)
                intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP") ' arrAnalytes(15, Count1)
                'intAnalyteIndex = arrAnalytes(3, Count1)'NOTE:  may be more than one analyteindex per group!!!
                intAnalyteID = tblAG.Rows(Count1 - 1).Item("ANALYTEID") ' arrAnalytes(2, Count1)
                strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                gstrAnal = strAnal

                'Legend
                'Dim arrAnalytes(15, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                ''10=UseIntStd, 11=IntStd, 12=MasterAssayID,13=IsCoadministeredCmpd,14=Original Analyte Description,15=Group

                If UseAnalyte(CStr(strDo)) Then
                Else
                    GoTo next1
                End If
                '* Different way of checking when samples are not assigned
                If boolNotAssignedSamples Then
                    bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                Else
                    bool = True
                End If


                'get an example AssayID from tblcalstdgroupassayid
                Dim rowsAllRuns() As DataRow

                strFAssayID = ""
                If boolUseGroups Then

                    If boolNotAssignedSamples Then

                        'tblBCStds1 = tblCalStdGroupsAcc

                        strF = "INTGROUP = " & intGroup
                        'don't need strFAssayID here
                        drows = tblBCStds1.Select(strF)

                        'get strFAssayID to possibly use in the future
                        rowsAllRuns = tblCalStdGroupAssayIDsAcc.Select(strF)
                        If rowsAllRuns.Length = 0 Then
                            GoTo end1
                        End If
                        strFAssayID = GetASSAYIDFilterIDCT(intGroup, False, True, intTableID)

                    Else

                        str1 = "ANALYTEID = " & intAnalyteID
                        drows = tblBCStds1.Select(str1)

                    End If

                Else
                    '20150811 Larry: Need to consider ANALYTEINDEX
                    '20160225 LEE: Not any more
                    str1 = "ANALYTEID = " & intAnalyteID & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1)
                    drows = tblBCStds1.Select(str1)

                End If

                If drows.Length = 0 Then
                    GoTo end1
                End If

                Dim strM As String
                If bool Then 'continue

                    intTCur = intTCur + 1

                    strM = "Creating " & strTName & " For " & strAnalC & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    '* This Analyte is good to go: Continue with the Report Preparation

                    intNRPos = 0
                    intNRUsed = 0

                    intRPIncr = 1
                    intRegrCt = 0

                    '* Page setup according to configuration, and Progress Update
                    str1 = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    Select Case intTableID
                        Case Is = 28
                            str2 = "Creating " & strTempInfo & " Summary of Back Calculated Standard Concentrations Table for " & strAnalC & "..."
                        Case Is <> 28
                            str2 = "Creating " & strTName & " for " & strAnalC & "..."
                            If intRunNum = 0 Then 'NOT assigned samples
                                'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                                Call InsertPageBreak(wd)
                            Else
                            End If
                            Call PageSetup(wd, str1) 'L=Landscape, P=Portrait
                    End Select
                    'strM = str2
                    'frmH.lblProgress.Text = str2
                    'frmH.Refresh()

                    ReDim arrBCStds(2, 100) '1=LevelNumber, 2=Concentration

                    'Legend
                    'Dim arrAnalytes(15, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                    '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                    ''10=UseIntStd, 11=IntStd, 12=MasterAssayID,13=IsCoadministeredCmpd,14=Original Analyte Description,15=Group

                    '

                    Dim tblTTc As DataTable
                    If boolEx Then 'if true, then is from assigned samples
                        Dim tblTTa As DataTable = drows.CopyToDataTable
                        ' Dim tblSR As System.Data.DataTable = dvT.ToTable("sr", True, "REGRESSIONTEXT")
                        Dim dvTTb As DataView = New DataView(tblTTa)
                        tblTTc = dvTTb.ToTable("TTc", True, strLevelNum, strNomConc)
                    Else
                        'here, drows is already uni
                        tblTTc = drows.CopyToDataTable
                    End If

                    Dim drowsBCS() As System.Data.DataRow
                    If boolUseGroups Then
                        If boolNotAssignedSamples Then
                            str1 = "INTGROUP = " & intGroup
                        Else
                            str1 = "ANALYTEID = " & intAnalyteID
                        End If
                    Else
                        '20150811 Larry: Need to consider ANALYTEINDEX
                        str1 = "ANALYTEID = " & intAnalyteID & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1)
                    End If

                    drowsBCS = tblBCStds1.Select(str1) 'this is used for Assigned Samples option

                    var1 = drowsBCS.Length 'debug


                    'int1 = drows.Length
                    int1 = tblTTc.Rows.Count
                    For Count2 = 0 To int1 - 1
                        If boolEx Then 'if true, then is from assigned samples
                        Else
                            'arrBCStds(1, Count2 + 1) = drows(Count2).Item("LevelNumber")
                            arrBCStds(1, Count2 + 1) = tblTTc.Rows(Count2).Item(strLevelNum) ' drows(Count2).Item(strLevelNum)
                        End If
                        'arrBCStds(2, Count2 + 1) = NZ(drows(Count2).Item("CONCENTRATION"), 0)
                        arrBCStds(2, Count2 + 1) = NZ(tblTTc.Rows(Count2).Item(strNomConc), 0) ' NZ(drows(Count2).Item(strNomConc), 0)
                    Next
                    ctCalibrStds = int1

                    Dim strRunID As String 'need this for later
                    Dim strBase As String
                    strBase = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND ANALYTEID = " & intAnalyteID

                    If intRunNum = 0 Then 'NOT assigned samples

                        If boolUseGroups Then
                            If boolIncludePSAE Then
                                str1 = "RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND ANALYTEID = " & intAnalyteID
                            Else
                                str1 = "RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND ANALYTEID = " & intAnalyteID
                            End If
                        Else
                            'If boolIncludePSAE Then
                            '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEID = " & intAnalyteID
                            'Else
                            '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEID = " & intAnalyteID
                            'End If
                            If boolIncludePSAE Then
                                str1 = "RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND ANALYTEID = " & intAnalyteID
                            Else
                                str1 = "RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND ANALYTEID = " & intAnalyteID
                            End If
                        End If

                        'ensure item is selected in anal run summary
                        strFFF = GetARSRuns(tblBCStdConcs1, intAnalyteID, "", False)
                        If Len(strFFF) = 0 Then
                        Else
                            strFFF = "(" & strFFF & ")"
                            str1 = str1 & " AND " & strFFF
                        End If

                    Else

                        'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & intAnalyteID
                        str1 = "ANALYTEID = " & intAnalyteID
                        str1 = str1 & " AND ("
                        int1 = rows.Length
                        intNumRuns = int1 'rows.Length
                        str2 = ""
                        For Count2 = 0 To int1 - 1
                            var1 = rows(Count2)
                            If Count2 = int1 - 1 Then
                                str2 = str2 & "RUNID = " & var1 & ")"
                            Else
                                str2 = str2 & "RUNID = " & var1 & " OR "
                            End If
                        Next
                        str1 = str1 & str2

                        'use the str2 portion for later use rowsassay
                        strFAID = "INTGROUP = " & intGroup
                        strFAID = strFAID & " AND (" & str2
                        'rowsassays

                    End If


                    'change drows to tblBCStdConcs1

                    strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC, ASSAYLEVEL ASC"
                    drows = tblBCStdConcs1.Select(str1, strS)

                    If intRunNum = 0 Then


                        '20161208 LEE: Must sync with Column B of analytical run summary

                        Dim dvT1 As DataView = New DataView(tblBCStdConcs1, str1, strS, DataViewRowState.CurrentRows)
                        Dim tblRID As DataTable = dvT1.ToTable("a", True, "RUNID")

                        '******

                        '20161208 LEE: The determination of intNumRegrWt and intNumRegr doesn't take into account AnalRunReview table column 'B'
                        'ensure item is selected in anal run summary

                        'Dim dv1 As System.Data.DataView
                        'dv1 = frmH.dgvAnalyticalRunSummary.DataSource
                        'Dim tblARS As System.Data.DataTable = dv1.ToTable
                        'Dim strFFF As String
                        'Dim intFFF As Short = 0
                        'Dim strFARS As String
                        'Dim rowsARS() As DataRow
                        'Dim intARS As Int16

                        'For Count2 = 0 To tblRID.Rows.Count - 1

                        '    var1 = tblRID.Rows(Count2).Item("RUNID")
                        '    'strFARS = "[Watson Run ID] = '" & var1 & "' AND Analyte_C = '" & strAnalC & "' AND BOOLINCLUDEREGR = " & True
                        '    strFARS = "[Watson Run ID] = '" & var1 & "' AND ANALYTEID = " & intAnalyteID & " AND BOOLINCLUDEREGR = " & True
                        '    Erase rowsARS
                        '    rowsARS = tblARS.Select(strFARS)
                        '    intARS = rowsARS.Length
                        '    If intARS = 0 Then
                        '    Else
                        '        intFFF = intFFF + 1
                        '        If intFFF = 1 Then
                        '            strFFF = "RUNID = " & var1
                        '        Else
                        '            strFFF = strFFF & " OR RUNID = " & var1
                        '        End If
                        '    End If
                        'Next


                        strFFF = GetARSRuns(tblRID, intAnalyteID, "", False)

                        If Len(strFFF) = 0 Then
                        Else
                            strFFF = "(" & strFFF & ")"
                            str1 = str1 & " AND " & strFFF
                        End If

                        '******

                        Erase drows
                        drows = tblBCStdConcs1.Select(str1, strS)

                    End If

                    int1 = drows.Length
                    intNumSamples = int1

                    Dim rowsLabels() As DataRow

                    ReDim arrBCStdConcs(10, intNumSamples + 100)
                    '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag, 5=AnalyteFlagPercent
                    '6=Hi, 7=Lo, 8=varNom, 9=minAnalFlag


                    'need to determine if Flags are used in this data
                    'first loop is through tblBCStdsAssayID to check each analytical run
                    Dim intZR As Short
                    intZR = 0

                    str1 = "INTGROUP = " & intGroup


skip1:

                    Count4 = 0
                    'For Count2 = 0 To int1 - 1 Step ctCalibrStds
                    Dim Level1, Level2
                    Dim intC As Short
                    intC = 0
                    Dim intRAID As Short = 0
                    Try

                        'redo this
                        'loop through assayids from tblCalStdGroupAssayIDs
                        Try
                            If boolNotAssignedSamples Then

                                str1 = "INTGROUP = " & intGroup
                                'ensure item is selected in anal run summary
                                strFFF = GetARSRuns(tblCalStdGroupAssayIDsAcc, intAnalyteID, "", False)

                                If Len(strFFF) = 0 Then
                                Else
                                    strFFF = "(" & strFFF & ")"
                                    str1 = str1 & " AND " & strFFF
                                End If

                                If BOOLINCLUDEDATE Then
                                    rowsAssays = tblCalStdGroupAssayIDsAcc.Select(str1, "RUNDATE ASC, RUNID ASC")
                                Else
                                    rowsAssays = tblCalStdGroupAssayIDsAcc.Select(str1, "RUNID ASC")
                                End If
                            Else
                                If BOOLINCLUDEDATE Then
                                    rowsAssays = tblCalStdGroupAssayIDsAcc.Select(strFAID, "RUNDATE ASC, RUNID ASC")
                                Else
                                    rowsAssays = tblCalStdGroupAssayIDsAcc.Select(strFAID, "RUNID ASC")
                                End If

                            End If
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try


                        Dim numLLOQ As Single
                        Dim numULOQ As Single
                        var1 = var1 'debug
                        For Count3 = 0 To rowsAssays.Length - 1 'this is each assay

                            'Legend
                            'arrBCStdConcs
                            '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag, 5=AnalyteFlagPercent
                            '6=Hi, 7=Lo, 8=varNom, 9=minAnalFlag

                            intAssayID = rowsAssays(Count3).Item("ASSAYID")
                            intAnalyteIndex = rowsAssays(Count3).Item("ANALYTEINDEX")
                            intRunID = rowsAssays(Count3).Item("RUNID")
                            numLLOQ = rowsAssays(Count3).Item("LLOQ")
                            numULOQ = rowsAssays(Count3).Item("ULOQ")

                            'strF = "ASSAYID = " & intAssayID & " AND ANALYTEINDEX = " & intAnalyteIndex & " AND ANALYTEID = " & intAnalyteID
                            strF = "ASSAYID = " & intAssayID & " AND ANALYTEID = " & intAnalyteID
                            Dim rowsNomConcs() As DataRow
                            Dim strNCF As String
                            Try
                                If boolNotAssignedSamples Then
                                    'rowsNomConcs = tblCalStdGroupsAcc.Select("INTGROUP = " & intGroup, "LEVELNUMBER ASC")
                                    rowsNomConcs = tblCalStdGroupsAcc.Select("INTGROUP = " & intGroup & " AND " & strF, "CONCENTRATION ASC")
                                    strNCF = "CONCENTRATION"
                                Else
                                    'get nomconc from tblassigned samples
                                    'Dim tblNC2 As System.Data.DataTable = dvNC2.ToTable("b", True, "ANALYTEID", "MASTERASSAYID", "ANALYTEINDEX", "NOMCONC", "STUDYID", "ASSAYID")
                                    rowsNomConcs = tblBCStds1.Select(strF, "NOMCONC")
                                    strNCF = "NOMCONC"
                                End If
                            Catch ex As Exception
                                var1 = ex.Message
                                var1 = var1
                            End Try

                            Dim intLevelNumber As Int16
                            Dim numFlagPercent As Single
                            Dim rowsLN() As DataRow

                            For Count4 = 0 To rowsNomConcs.Length - 1 'this is nomconc loop

                                intC = intC + 1

                                Try
                                    If boolNotAssignedSamples Then
                                        intLevelNumber = rowsNomConcs(Count4).Item("LEVELNUMBER")
                                    Else
                                        'get levelnumber from tblCalStdGroupsAll
                                        var1 = rowsNomConcs(Count4).Item("ASSAYID")
                                        var2 = rowsNomConcs(Count4).Item("NOMCONC")
                                        str1 = "INTGROUP = " & intGroup & " AND CONCENTRATION = " & var2
                                        rowsLN = tblCalStdGroupsAll.Select(str1)
                                        intLevelNumber = rowsLN(0).Item("LEVELNUMBER")
                                    End If
                                Catch ex As Exception
                                    var1 = ex.Message
                                    var1 = var1
                                End Try

                                arrBCStdConcs(1, intC) = intLevelNumber

                                'Legend
                                'arrBCStdConcs
                                '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag, 5=AnalyteFlagPercent
                                '6=Hi, 7=Lo, 8=varNom, 9=minAnalFlag

                                'need to get eliminated flag from tblbcstdconcs
                                strF1 = strF & " AND ASSAYLEVEL = " & intLevelNumber 'Assumes LEVELNUMBER = ASSAYLEVEL
                                Dim rowsAF() As DataRow = tblBCStdConcs1.Select(strF1)
                                If rowsAF.Length = 0 Then
                                    arrBCStdConcs(2, intC) = "NI"
                                    arrBCStdConcs(4, intC) = "Y"
                                Else
                                    strAnalayteFlag = rowsAF(0).Item("ELIMINATEDFLAG")
                                    num1 = rowsNomConcs(Count4).Item(strNCF)
                                    If boolLUseSigFigs Then
                                        num2 = SigFigOrDec(CDec(num1), LSigFig, False)
                                    Else
                                        num2 = RoundToDecimalRAFZ(CDec(num1), LSigFig)
                                    End If
                                    arrBCStdConcs(2, intC) = num2
                                    arrBCStdConcs(4, intC) = strAnalayteFlag
                                End If

                                arrBCStdConcs(3, intC) = intRunID

                                num1 = rowsNomConcs(Count4).Item(strNCF)
                                If boolLUseSigFigs Then
                                    numNomConc = SigFigOrDec(CDec(num1), LSigFig, False)
                                Else
                                    numNomConc = RoundToDecimalRAFZ(CDec(num1), LSigFig)
                                End If
                                arrBCStdConcs(8, intC) = numNomConc

                                Try
                                    If boolNotAssignedSamples Then
                                        numFlagPercent = NZ(rowsNomConcs(Count4).Item("ANALYTEFLAGPERCENT"), 15)
                                    Else
                                        numFlagPercent = NZ(rowsLN(0).Item("ANALYTEFLAGPERCENT"), 15)
                                    End If

                                Catch ex As Exception
                                    var1 = ex.Message
                                    var1 = var1
                                End Try
                                If intRunNum = 0 Then 'NOT assigned samples

                                    arrBCStdConcs(5, intC) = numFlagPercent
                                    arrBCStdConcs(9, intC) = numFlagPercent

                                    Call SetHighAndLowCriteria(CDbl(numNomConc), CDbl(numFlagPercent), CDbl(numFlagPercent), hi, lo)
                                    'Call SetHighAndLowCriteria(numNomConc, numFlagPercent, numFlagPercent, hi, lo)
                                    arrBCStdConcs(6, intC) = hi
                                    arrBCStdConcs(7, intC) = lo

                                    'var1 = numNomConc * numFlagPercent / 100
                                    'hi = numNomConc + var1
                                    'lo = numNomConc - var1

                                    'If boolLUseSigFigs Then
                                    '    arrBCStdConcs(6, intC) = SigFigOrDec(hi, LSigFig, False)
                                    'Else
                                    '    arrBCStdConcs(6, intC) = RoundToDecimalRAFZ(hi, LSigFig)
                                    'End If

                                    'If boolLUseSigFigs Then
                                    '    arrBCStdConcs(7, intC) = SigFigOrDec(lo, LSigFig, False)
                                    'Else
                                    '    arrBCStdConcs(7, intC) = RoundToDecimalRAFZ(lo, LSigFig)
                                    'End If

                                Else
                                    'Try
                                    '    boolRAID = False
                                    '    var4 = numNomConc 'drowsBCS(Count4 - 1).Item("NOMCONC")
                                    '    For Count5 = 0 To rowsAssID.Length - 1
                                    '        var3 = rowsAssID(Count5).Item("CONCENTRATION")
                                    '        If var3 = var4 Then
                                    '            boolRAID = True
                                    '            Exit For
                                    '        End If
                                    '    Next
                                    '    If boolRAID Then
                                    '        var1 = rowsAssID(Count5).Item("ANALYTEFLAGPERCENT")
                                    '        var2 = NZ(rowsAssID(Count5).Item("ANALYTEFLAGPERCENT"), 15) 'debugging
                                    '    Else
                                    '        var2 = 15 'System.DBNull.Value
                                    '    End If

                                    'Catch ex As Exception
                                    '    var2 = 15 'System.DBNull.Value
                                    'End Try

                                    'v1 = var2
                                    'v2 = var2

                                    v1 = numFlagPercent
                                    v2 = numFlagPercent

                                    Try
                                        If gAllowGuWuAccCrit And LAllowGuWuAccCrit Then
                                            var1 = NZ(rowsAF(0).Item("BOOLUSEGUWUACCCRIT"), 0)
                                            If var1 = 0 Then
                                            Else
                                                v1 = NZ(rowsAF(0).Item("NUMMAXACCCRIT"), 15)
                                                v2 = NZ(rowsAF(0).Item("NUMMINACCCRIT"), 15)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        var1 = ex.Message
                                        var1 = var1
                                    End Try


                                    arrBCStdConcs(5, intC) = v1 'max percentflag
                                    arrBCStdConcs(9, intC) = v2 'min percentflag

                                    Call SetHighAndLowCriteria(CDbl(numNomConc), CDbl(v1), CDbl(v2), hi, lo)

                                    arrBCStdConcs(6, intC) = hi
                                    arrBCStdConcs(7, intC) = lo

                                    'hi = var4 * v1 / 100
                                    'lo = var4 * v1 / 100

                                    'If boolLUseSigFigs Then
                                    '    arrBCStdConcs(6, intC) = SigFigOrDec(hi, LSigFig, False)
                                    'Else
                                    '    arrBCStdConcs(6, intC) = RoundToDecimalRAFZ(hi, LSigFig)
                                    'End If

                                    'If boolLUseSigFigs Then
                                    '    arrBCStdConcs(7, intC) = SigFigOrDec(lo, LSigFig, False)
                                    'Else
                                    '    arrBCStdConcs(7, intC) = RoundToDecimalRAFZ(lo, LSigFig)
                                    'End If

                                    arrBCStdConcs(8, intC) = numNomConc


                                    'Legend
                                    'arrBCStdConcs
                                    '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag, 5=AnalyteFlagPercent
                                    '6=Hi, 7=Lo, 8=varNom, 9=minAnalFlag

                                    '******


                                End If

                                'Legend
                                'arrBCStdConcs
                                '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag, 5=AnalyteFlagPercent
                                '6=Hi, 7=Lo, 8=varNom, 9=minAnalFlag

                                '*****
                                var1 = var1 'debug

                            Next Count4

                            var1 = var1 'debug

                        Next Count3

                        var1 = var1 'debug
                        'intNR = intZR

                    Catch ex As Exception
                        'MsgBox(Count3 & " out of " & int1 - 1)
                        var1 = var1 'debug
                    End Try

                    'determine number of calibr std reps
                    Dim tblBB As DataTable

                    'tblBCStdConcs1 doesn't have GROUP column if UseGroups = false

                    Dim intBB As Int64
                    If boolNotAssignedSamples Then
                        str1 = "INTGROUP = " & intGroup
                    Else
                        str1 = "ANALYTEID = " & intAnalyteID
                    End If
                    Dim rowsAA() As DataRow = tblBCStds1.Select(str1)

                    If boolUseGroups Then
                        tblBB = rowsAA.CopyToDataTable
                    Else
                        'get unique assayID's from rowsAA

                        Dim dvAA As DataView = New DataView(rowsAA.CopyToDataTable)
                        Try
                            tblBB = dvAA.ToTable("AA", True, "ASSAYID", "RUNID")
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try


                    End If

                    intCStdReps = 0
                    For Count2 = 0 To tblBB.Rows.Count - 1
                        intBB = NZ(tblBB.Rows(Count2).Item("ASSAYID"), 0)
                        intRunID = NZ(tblBB.Rows(Count2).Item("RUNID"), 0)
                        If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                            strF = "RUNID = " & intRunID & " AND RUNTYPEID > 0 AND ANALYTEID = " & intAnalyteID
                        Else
                            strF = "RUNID = " & intRunID & " AND RUNTYPEID <> 3 AND ANALYTEID = " & intAnalyteID
                            'strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 & " AND RUNID = " & var3 & " AND RUNTYPEID <> 3 AND ANALYTEID = " & var4
                        End If
                        'ignore PSAE
                        strF = "RUNID = " & intRunID & " AND RUNTYPEID <> 3 AND ANALYTEID = " & intAnalyteID
                        rowsCSR = tblBCStdConcs1.Select(strF)
                        int1 = rowsCSR.Length / ctCalibrStds
                        If int1 > intCStdReps Then
                            intCStdReps = int1
                        End If
                    Next

                    '*****

                    'find number of regression parameters
                    'Dim dvT As System.Data.DataView = New DataView(tblRegCon)
                    Dim dvT As System.Data.DataView = New DataView(tblCalStdGroupAssayIDsAcc)
                    'Dim dvT As System.Data.DataView = New DataView(tblRegConAll)

                    'If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                    '    str1 = "RUNTYPEID > 0 AND ANALYTEID = " & intAnalyteID
                    'Else
                    '    str1 = "RUNTYPEID <> 3 AND ANALYTEID = " & intAnalyteID
                    'End If

                    If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                        str1 = "RUNTYPEID > 0 AND INTGROUP = " & intGroup
                    Else
                        str1 = "RUNTYPEID <> 3 AND INTGROUP = " & intGroup
                    End If

                    'str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                    dvT.RowFilter = str1

                    'loop through dvt to find regr information

                    intRP = 0
                    intNumRegr = 0
                    ReDim arrRegrType(2, 100)
                    '1=Regr Type, 2 = weighting
                    'Note: 2 isn't used

                    For Count2 = 0 To dvT.Count - 1
                        intAssayID = dvT(Count2).Item("ASSAYID")
                        intAnalyteIndex = dvT(Count2).Item("ANALYTEINDEX")

                        'strF = "ASSAYID = " & intAssayID & " AND ANALYTEINDEX = " & intAnalyteIndex
                        strF = "ASSAYID = " & intAssayID & " AND ANALYTEID = " & intAnalyteID
                        Dim rowsRC() As DataRow = tblRegCon.Select(strF)
                        int1 = rowsRC.Length
                        If int1 > intRP Then
                            intRP = int1
                        End If

                        'look for different regressions
                        If rowsRC.Length = 0 Then
                        Else
                            var1 = NZ(rowsRC(0).Item("REGRESSIONTEXT"), "NA")
                            If StrComp(var1, "NA", CompareMethod.Text) = 0 Then
                            Else
                                If intNumRegr = 0 Then
                                    intNumRegr = intNumRegr + 1
                                    arrRegrType(1, intNumRegr) = var1
                                Else
                                    boolHit = False
                                    For Count3 = 1 To intNumRegr
                                        var2 = NZ(rowsRC(0).Item("REGRESSIONTEXT"), "NA")
                                        If StrComp(var2, "NA", CompareMethod.Text) = 0 Then
                                        Else
                                            If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                                                boolHit = True
                                                Exit For
                                            End If
                                        End If
                                    Next Count3
                                    If boolHit Then
                                    Else
                                        intNumRegr = intNumRegr + 1
                                        arrRegrType(1, intNumRegr) = var1
                                    End If
                                End If
                            End If

                        End If

                    Next Count2


                    ReDim Preserve arrRegrType(2, intNumRegr)
                    If intNumRegr = 1 Then
                        boolSRegr = True
                    Else
                        boolSRegr = False
                    End If


                    'intRP = dvT.Count

                    'var1 = dvT.Count 'debugging

                    'Dim tblT As System.Data.DataTable = dvT.ToTable("a", True, "REGRESSIONPARAMETERID")
                    'intRP = tblT.Rows.Count

                    ''determine if there is more than one regr type
                    'Dim tblSR As System.Data.DataTable = dvT.ToTable("sr", True, "REGRESSIONTEXT")
                    'intNumRegr = tblSR.Rows.Count

                    'ReDim arrRegrType(2, intNumRegr)
                    'If intNumRegr = 1 Then
                    '    boolSRegr = True
                    'Else
                    '    boolSRegr = False
                    'End If
                    'For Count2 = 1 To intNumRegr
                    '    var1 = tblSR.Rows(Count2 - 1).Item("REGRESSIONTEXT") 'Regression type
                    '    arrRegrType(1, Count2) = tblSR.Rows(Count2 - 1).Item("REGRESSIONTEXT")
                    'Next

                    'find number of table rows
                    Dim tblRunID As System.Data.DataTable
                    intTRows = 0
                    If boolEx Then 'comes from assigned StdCalibr
                        'HERE
                        'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, 
                        '" & strAnaRunPeak & ".RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, 
                        'ANARUNANALYTERESULTS.ANALYTEINDEX, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, 
                        'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, 
                        'ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS

                        'to determine inttemprows, find unique runid
                        str1 = "ANALYTEID = " & intAnalyteID
                        str2 = "ANALYTEID ASC"
                        Dim dvTT As System.Data.DataView = New DataView(tblBCStdConcs1, str1, str2, DataViewRowState.CurrentRows)
                        tblRunID = dvTT.ToTable("tt", True, "RUNID")
                        intTRows = tblRunID.Rows.Count
                        intNumRuns = intTRows
                        For Count2 = 0 To intNumRuns - 1 'debugging
                            var1 = tblRunID.Rows(Count2).Item("RunID")
                            var2 = var1
                        Next

                    Else

                        'tblRunID = dvT.ToTable("b", True, "RUNID")
                        'intTRows = tblRunID.Rows.Count
                        'intNumRuns = intTRows
                        'inttemprows tblRunID intTRows 'arrAnalytes(7, Count1) '# of accepted runs
                        tblRunID = rowsAssays.CopyToDataTable
                        intNumRuns = rowsAssays.Length
                        intTRows = intNumRuns

                    End If

                    inttemprows = intTRows
                    intRegrCt = intTRows
                    ReDim arrRegCon(intRP + 2, intRegrCt)
                    ReDim arrRunID(intRegrCt)

                    'determine number of weightings
                    Dim strW As String

                    'get ULOQUnits
                    dv = frmH.dgvWatsonAnalRef.DataSource
                    int1 = FindRowDV("ULOQ Units", dv)
                    strConcUnits = dv.Item(int1).Item(arrAnalytes(1, Count1))

                    int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strConcUnits = str1
                    End If

                    strW = "NA"
                    For Count2 = 0 To intNumRuns - 1
                        If Count2 = 0 Then
                            var3 = tblRunID.Rows(Count2).Item("RUNID")
                            'strW = GetWtRegCon(CInt(var3), intAnalyteID)
                            var4 = GetWtRegCon(CInt(var3), intAnalyteID)
                            If StrComp(var4, "NA", CompareMethod.Text) = 0 Then
                            Else
                                strW = var4
                            End If
                            var1 = var4
                            var2 = var1
                        Else
                            var2 = var1
                            var3 = tblRunID.Rows(Count2).Item("RUNID")
                            'strW = GetWtRegCon(CInt(var3), intAnalyteID)
                            var4 = GetWtRegCon(CInt(var3), intAnalyteID)
                            If StrComp(var4, "NA", CompareMethod.Text) = 0 Then
                            Else
                                strW = var4
                            End If
                            var1 = var4
                        End If


                        If StrComp(var1, var2, CompareMethod.Text) = 0 Or StrComp(var1, "NA", CompareMethod.Text) <> 0 Or StrComp(var2, "NA", CompareMethod.Text) <> 0 Then
                        Else
                            boolSWt = False
                            Exit For
                        End If

                        If Count2 = intNumRuns - 1 Then
                            strWeighting = strW
                        End If
                    Next
                    '*****

                    'calculate # of table rows
                    int1 = 0
                    int1 = int1 + 3 'for header rows
                    int1 = int1 + 1 'blank row
                    'If intCStdReps = 1 Then
                    '    'int1 = (inttemprows * (intCStdReps)) - 1 + 10
                    '    int1 = int1 + (inttemprows * (intCStdReps)) 'for data rows
                    '    int1 = int1 + 1 'for blank row
                    'Else 'separate sets with a blank row
                    '    'int1 = (inttemprows * (intCStdReps + 1)) - 1 + 10
                    '    int1 = int1 + (inttemprows * (intCStdReps + 1)) 'includes blank row between data
                    'End If

                    int1 = int1 + (inttemprows * (intCStdReps + 1)) 'includes blank row between data

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    int1 = int1 + intCSN

                    If boolCSREPORTACCVALUES Then
                        'show stats for accepted values only

                    Else
                        If intExp > 0 Then

                            int1 = int1 + 2 'for two extra stats section lables
                            'int1 = int1 + (2 * 5) + 1 'for 2 stats sections separated by a space

                            'Increment for Statistics Sections
                            int1 = int1 + intCSN
                            If intCSN > 0 Then
                                int1 = int1 + 1
                            End If

                        End If
                    End If

                    If BOOLINCLUDEDATE Then
                        'int1 = int1 + 1
                    End If


                    int3 = int1

                    '***determine if regression section is to be added

                    ''record strregressiontype
                    ''don't need to do this here
                    'int1 = FindRow("Regression", tblWatsonAnalRefTable, "Item")
                    'int2 = FindRow("Weighting", tblWatsonAnalRefTable, "Item")
                    'strRegressionType = tblWatsonAnalRefTable.Rows(int1).Item(Count1)
                    ''strWeighting = tblWatsonAnalRefTable.Rows(int2).Item(Count1) 'get this earlier

                    Count2 = 0
                    '1=RUNID, 2=AnalyteIndex, 3=REGRESSIONPARAMETERID(1=Slope, 2=YInt, 3=R2),4=PARAMETERVALUE
                    '1=RUNID,  2=Slope, 3=YInt, 4=R2
                    '1=RUNID, intRP parameters, intRP+2=R2


                    strF = str1

                    'LEGEND: arrRegCon(intRP + 2, intRegrCt)
                    For Count3 = 1 To intRegrCt
                        var1 = tblRunID.Rows(Count3 - 1).Item("RUNID")
                        'If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND ANALYTEID = " & intAnalyteID
                        'Else
                        '    strF = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3  AND RUNID = " & var1 & " AND ANALYTEID = " & intAnalyteID
                        'End If

                        If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                            strF = "STUDYID = " & wStudyID & " AND RUNTYPEID > 0 AND RUNID = " & var1 & " AND ANALYTEID = " & intAnalyteID
                        Else
                            strF = "STUDYID = " & wStudyID & " AND RUNTYPEID <> 3  AND RUNID = " & var1 & " AND ANALYTEID = " & intAnalyteID
                        End If

                        'use RegCon for this action
                        'drows = tblRegConAll.Select(strF, "REGRESSIONPARAMETERID ASC")
                        drows = tblRegCon.Select(strF, "REGRESSIONPARAMETERID ASC")
                        int1 = drows.Length
                        If int1 = 0 Then 'analytical run has not been accepted
                            arrRegCon(1, Count3) = NZ(var1, "NA")
                            arrRegCon(2, Count3) = "NA" '0
                            'var3 = drows(0).Item("RSQUARED")
                            var3 = "NA" ' 0 'drows(0).Item("RSQUARED")
                            arrRegCon(intRP + 2, Count3) = var3
                            arrRunID(Count3) = var1 ' drows(0).Item("RUNID")
                        Else
                            arrRegCon(1, Count3) = var1
                            For Count4 = 0 To int1 - 1
                                var2 = NZ(drows(Count4).Item("PARAMETERVALUE"), "NA")
                                If IsNumeric(var2) Then
                                    If boolLUseSigFigsRegr Then
                                        var2 = SigFigOrDec(NZ(drows(Count4).Item("PARAMETERVALUE"), 0), LRegrSigFigs, False)
                                        arrRegCon(Count4 + 2, Count3) = Format(var2, GetScNot(LRegrSigFigs))
                                    Else
                                        var2 = RoundToDecimalRAFZ(NZ(drows(Count4).Item("PARAMETERVALUE"), 0), LRegrSigFigs)
                                        arrRegCon(Count4 + 2, Count3) = Format(var2, GetScNot(LRegrSigFigs))
                                    End If
                                Else
                                    arrRegCon(Count4 + 2, Count3) = var2
                                End If


                                ' arrRegCon(Count4 + 2, Count3) = Format(var2, GetScNot(LRegrSigFigs))
                            Next
                            'var3 = drows(0).Item("RSQUARED")
                            var3 = NZ(drows(0).Item("RSQUARED"), "NA")
                            'If boolLUseSigFigsRegr Then
                            '    arrRegCon(intRP + 2, Count3) = CStr(SigFigOrDecString(NZ(var3, 0), LR2SigFigs, False))
                            'Else
                            '    arrRegCon(intRP + 2, Count3) = Format(RoundToDecimalRAFZ(NZ(var3, 0), LR2SigFigs), GetRegrDecStr(LR2SigFigs))
                            'End If
                            If IsNumeric(var3) Then
                                If boolLUseSigFigsRegr Then
                                    'arrRegCon(intRP + 2, Count3) = CStr(SigFigOrDecString(NZ(var3, 0), LR2SigFigs, False))
                                    var2 = SigFigOrDecString(NZ(var3, 0), LR2SigFigs, False)
                                    arrRegCon(intRP + 2, Count3) = Format(CDec(var2), GetScNot(LR2SigFigs))
                                Else
                                    'arrRegCon(intRP + 2, Count3) = Format(RoundToDecimalRAFZ(NZ(var3, 0), LR2SigFigs), GetRegrDecStr(LR2SigFigs))
                                    var2 = RoundToDecimalRAFZ(NZ(var3, 0), LR2SigFigs)
                                    arrRegCon(intRP + 2, Count3) = Format(CDec(var2), GetScNot(LR2SigFigs))
                                End If
                            Else
                                arrRegCon(intRP + 2, Count3) = "NA"
                            End If


                            arrRunID(Count3) = drows(0).Item("RUNID")
                        End If

                    Next

                    Dim numCols As Short
                    If boolSTATSDIFFCOL Then
                        If boolSTATSREGR Then
                            If boolSRegr And boolSWt Then
                                numCols = (ctCalibrStds * 2) + 1 + intRP + 1 ' + 1 '1 more for regression 
                            ElseIf boolSRegr Or boolSWt Then
                                numCols = (ctCalibrStds * 2) + 1 + intRP + 1 + 1 '1 more for regression or weighting 
                            Else
                                numCols = (ctCalibrStds * 2) + 1 + intRP + 1 + 2 '2 more for regression and weighting
                            End If
                        Else
                            numCols = (ctCalibrStds * 2) + 1
                        End If
                    Else
                        If boolSTATSREGR Then
                            If boolSRegr And boolSWt Then
                                numCols = (ctCalibrStds * 1) + 1 + intRP + 1 ' + 1 '1 more for regression 
                            ElseIf boolSRegr Or boolSWt Then
                                numCols = (ctCalibrStds * 1) + 1 + intRP + 1 + 1 '1 more for regression or weighting
                            Else
                                numCols = (ctCalibrStds * 1) + 1 + intRP + 1 + 2 '2 more for regression and weighting
                            End If
                            'numCols = ctCalibrStds + 1 + intRP + 1 ' + 2 '2 more for regression and weighting
                        Else
                            numCols = ctCalibrStds + 1
                        End If
                    End If

                    wrdSelection = wd.Selection()

                    '''''''wdd.visible = True


                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=int3, NumColumns:=numCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False

                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        .Selection.Rows.AllowBreakAcrossPages = False

                        'If boolSTATSMEAN Then
                        '    .Selection.font.size = 10
                        'End If

                        removeBorderButLeaveTopAndBottom(wd)
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        '.Selection.Font.Size = 11


                        If boolPlaceHolder Then

                            .Selection.Tables.Item(1).Select()
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone


                            'strA = arrAnalytes(14, Count1)
                            If gNumMatrix = 1 Then
                                strA = strAnalC
                            Else
                                strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                            End If
                            'No. Now just send strAnal
                            strA = strAnal
                            strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                            Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                            'Note: strTName is byRef and will return Table, number, caption, label

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

                        .Selection.Tables.Item(1).Cell(1, 2).Select()

                        ''wd.visible = True'debugging

                        If boolSTATSREGR Then

                            '''''''''''wdd.visible = True

                            'enter calibration curve params
                            .Selection.Tables.Item(1).Cell(1, 1 + (ctCalibrStds * int11) + 1).Select()
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try
                            '.Selection.Font.Size = 11
                            .Selection.Font.Bold = False
                            .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                            .Selection.TypeText(Text:="Calibration Curve Parameters")

                            .Selection.Tables.Item(1).Cell(1, 2).Select()
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=(ctCalibrStds * int11) - 1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try
                            '.Selection.Font.Size = 11
                            .Selection.Font.Bold = False
                            '.Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                            If LboolNomConcParen Then
                                .Selection.TypeText(Text:="Nominal Concentrations") ' (" & strConcUnits & ")")
                            Else
                                .Selection.TypeText(Text:="Nominal Concentrations")
                            End If

                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        Else

                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                            Try
                                .Selection.Cells.Merge()
                            Catch ex As Exception

                            End Try

                            '.Selection.Font.Size = 11
                            .Selection.Font.Bold = False
                            If LboolNomConcParen Then
                                .Selection.TypeText(Text:="Nominal Concentrations") ' (" & strConcUnits & ")")
                            Else
                                .Selection.TypeText(Text:="Nominal Concentrations")
                            End If
                            '.Selection.TypeText(Text:="Nominal Concentrations")
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        End If

                        'border top and bottom of range
                        '.Selection.Tables.item(1).Cell(1, 1).Select()
                        '.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle
                        '.Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleSingle


                        'underline
                        .Selection.Tables.Item(1).Cell(1, 1).Select()
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        '.Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)
                        Dim intHRow As Short
                        If BOOLINCLUDEDATE Then
                            intHRow = 3 '4
                        Else
                            intHRow = 3
                        End If

                        'underline
                        .Selection.Tables.Item(1).Cell(intHRow, 1).Select()
                        '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=ctCalibrStds, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.Borders.item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        .Selection.Tables.Item(1).Cell(intHRow - 1, 2).Select()
                        Dim intLabelRow As Short
                        intLabelRow = intHRow - 1

                        'record column nomconc headings


                        '20150811 Larry: Need to consider ANALYTEINDEX
                        'strF = "ANALYTEID = " & intAnalyteID & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1)

                        'strS = "Concentration ASC"

                        strS = strNomConc & " ASC"
                        If boolNotAssignedSamples Then
                            strF = "INTGROUP = " & intGroup
                        Else
                            strF = "ANALYTEID = " & intAnalyteID
                        End If
                        Dim dvN1 As System.Data.DataView = New DataView(tblBCStds1, strF, strS, DataViewRowState.CurrentRows)
                        'Dim tblNomConc As System.Data.DataTable = dvN1.ToTable("aaa", True, "Concentration", "LEVELNUMBER")

                        '20150723 Larry: tblBCStds1 doesn't have ASSAYID if tblBCStds1 IS NOT from assigned samples
                        'tblNomConcWithAssayID results in rows with NULL for ASSAYID
                        'Must use 'tblBCStdsAssayID instead
                        'Dim dvN1A As System.Data.DataView = New DataView(tblBCStdsAssayID, strF, "CONCENTRATION ASC", DataViewRowState.CurrentRows)

                        strF = "INTGROUP = " & intGroup
                        Dim dvN1A As System.Data.DataView = New DataView(tblCalStdGroupsAcc, strF, "CONCENTRATION ASC", DataViewRowState.CurrentRows)

                        Dim tblNomConcWithAssayID As System.Data.DataTable
                        If boolNotAssignedSamples Then
                            tblNomConcWithAssayID = dvN1A.ToTable("aaa", True, strNomConc, strLevelNum, "RUNID")
                        Else
                            tblNomConcWithAssayID = dvN1.ToTable("aaa", True, strNomConc, strLevelNum, "RUNID")
                        End If
                        'Dim tblNomConcWithAssayID As System.Data.DataTable = dvN1A.ToTable("aaa", True, strNomConc, strLevelNum, "ASSAYID")
                        Dim tblNomConc As System.Data.DataTable = dvN1.ToTable("aaa", True, strNomConc, strLevelNum)

                        ctCalibrStds = tblNomConc.Rows.Count

                        ''debug
                        'var1 = ""
                        'For Count3 = 0 To tblBCStds1.Columns.Count - 1
                        '    var2 = tblBCStds1.Columns(Count3).ColumnName
                        '    var1 = var1 & ";" & var2
                        'Next
                        '''''''console.writeline(var1)
                        'For Count2 = 0 To dvN1.Count - 1
                        '    var1 = ""
                        '    For Count3 = 0 To tblBCStds1.Columns.Count - 1
                        '        var2 = dvN1(Count2).Item(Count3)
                        '        var1 = var1 & ";" & var2
                        '    Next
                        '    ''''''console.writeline(var1)
                        'Next

                        For Count2 = 0 To ctCalibrStds - 1

                            var1 = tblNomConc.Rows(Count2).Item(strNomConc)
                            If IsNumeric(var1) Then
                                If boolLUseSigFigs Then
                                    str1 = CStr(DisplayNum(SigFigOrDec(var1, LSigFig, False), LSigFig, False))
                                Else
                                    str1 = CStr(Format(var1, GetRegrDecStr(LSigFig)))
                                End If
                            Else
                                If IsDBNull(var1) Then
                                    str1 = "NA"
                                Else
                                    str1 = CStr(var1)
                                End If

                            End If
                            str3 = str1


                            ''20150723 Larry:  This isn't working for study MethVal
                            If BOOLINCLUDEWATSONLABELS Then

                                'must do two steps

                                Dim intColW As Short
                                Dim intFW As Short
                                Dim rowsFW() As DataRow
                                Dim rowsFF() As DataRow
                                Dim strFF As String
                                Dim varLN
                                Dim strSort As String

                                int1 = tblBCStdsAssayID.Rows.Count 'debug

                                'Find first assayID that corresponds to this level
                                'Larry ?: AssayID is integer, not text
                                'strAssayID = ""
                                Dim thisNomConc As Single = NZ(tblNomConc.Rows(Count2).Item(strNomConc), -1)
                                Dim intThisNomConcAssayID As Int32

                                intAssayID = 0
                                For Count3 = 0 To tblNomConcWithAssayID.Rows.Count - 1
                                    'Collect the first AssayID for this Nominal Concentration. This will be used to find the label for this level.
                                    'In the case where we have two labels for the same nominal concentration in the same table, this won't work.

                                    'However, we could fix it fairly easily by just going through all the AssayIDs in the tblBCStdAssayID table
                                    'for a certain concentration, and looking at the label for each of the AssayIDs for that concentration,
                                    'and concatenating the labels if there were multiple lables for the same concentration on the same table.
                                    '
                                    '20150723 Larry:  This isn't working for study MethVal
                                    var1 = NZ(tblNomConcWithAssayID.Rows(Count3).Item(strNomConc), -2)
                                    var2 = NZ(tblNomConcWithAssayID.Rows(Count3).Item("ASSAYID"), -3)
                                    If (thisNomConc = var1 And var2 <> -3) Then
                                        intThisNomConcAssayID = NZ(tblNomConcWithAssayID.Rows(Count3).Item("ASSAYID"), -3)
                                        intAssayID = intThisNomConcAssayID
                                        Exit For
                                    End If
                                Next

                                strFF = "CONCENTRATION = " & CSng(var1) & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYID = " & intAssayID
                                strSort = "LEVELNUMBER ASC"
                                rowsFF = tblBCStdsAssayID.Select(strFF, strSort)
                                If rowsFF.Length = 0 Then

                                    '20150718 Larry: For some reason var1 = 0.1 returns 0 hits
                                    '20150720 NIck: It's because it is stored as 0.10000000000000002, which shows up as
                                    '"0" when cast to a double (which you do below).  However, it doesn't show up as 0 when
                                    'the SELECT command is used.

                                    'look further
                                    Dim rowsFFF() As DataRow
                                    Dim c1 As Single
                                    strFFF = "ANALYTEID = " & intAnalyteID
                                    rowsFFF = tblBCStdsAssayID.Select(strFFF, strSort)
                                    str1 = "NA"
                                    For Count3 = 0 To rowsFFF.Length - 1
                                        c1 = rowsFFF(Count3).Item("CONCENTRATION")
                                        If c1 = CSng(var1) Then
                                            var2 = rowsFFF(Count3).Item("LEVELNUMBER")
                                            strF = "LEVELNUMBER = " & var2 & " AND KNOWNTYPE = 'STANDARD'" & " AND ASSAYID = " & intAssayID
                                            rowsFW = tblAssayLabels.Select(strF)
                                            If rowsFW.Length = 0 Then
                                                str1 = "NA"
                                            Else
                                                str1 = NZ(rowsFW(0).Item("ID"), "NA")
                                            End If
                                            Exit For
                                        End If
                                    Next
                                Else
                                    'Dim tblNomConc As System.Data.DataTable = dvN1.ToTable("aaa", True, strNomConc, strLevelNum)
                                    var2 = rowsFF(0).Item("LEVELNUMBER")
                                    strF = "LEVELNUMBER = " & var2 & " AND KNOWNTYPE = 'STANDARD'" & " AND ASSAYID = " & intAssayID
                                    rowsFW = tblAssayLabels.Select(strF)
                                    If rowsFW.Length = 0 Then
                                        str1 = "NA"
                                    Else
                                        str1 = NZ(rowsFW(0).Item("ID"), "NA")
                                    End If
                                End If

                                str3 = str1 & ChrW(10) & str3

                            End If 'BOOLINCLUDEWATSONLABELS

                            ''*****
                            .Selection.TypeText(Text:=str3)
                            '.Selection.TypeText(Text:=CStr(SigFigOrDecString(CDbl(arrBCStds(2, Count2)), LSigFig, False)))
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            If boolSTATSDIFFCOL Then
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If
                        Next

                        ''''wdd.visible = True


                        'assay levels come from tblCalStdGroupAssayIDs
                        'filtered for INTGROUP in rowsAssays, sorted by RUNDATE


                        'record assaylevels in tblBCStdConcs
                        strF = "ANALYTEID = " & intAnalyteID
                        strS = "ASSAYLEVEL ASC"
                        Dim dvN2 As System.Data.DataView = New DataView(tblBCStdConcs1, strF, strS, DataViewRowState.CurrentRows)

                        ''debug
                        ''tblBCStdsAssayID
                        'var1 = ""
                        'For Count3 = 0 To tblBCStdConcs1.Columns.Count - 1
                        '    var2 = tblBCStdConcs1.Columns(Count3).ColumnName
                        '    var1 = var1 & ";" & var2
                        'Next
                        '''''''console.writeline(var1)
                        '''''''console.writeline("Start tblBCStdConcsq")

                        'For Count2 = 0 To tblBCStdConcs1.Rows.Count - 1
                        '    var1 = ""
                        '    For Count3 = 0 To tblBCStdConcs1.Columns.Count - 1
                        '        var2 = tblBCStdConcs1.Rows(Count2).Item(Count3)
                        '        var1 = var1 & ";" & var2
                        '    Next
                        '    ''''''console.writeline(var1)
                        'Next
                        '''''''console.writeline("Start dvN1")
                        'For Count2 = 0 To dvN2.Count - 1
                        '    var1 = ""
                        '    For Count3 = 0 To tblBCStdConcs1.Columns.Count - 1
                        '        var2 = dvN2(Count2).Item(Count3)
                        '        var1 = var1 & ";" & var2
                        '    Next
                        '    ''''''console.writeline(var1)
                        'Next

                        'Dim tblAssayLevels As System.Data.DataTable = dvN2.ToTable("bbb", True, "ASSAYLEVEL")
                        Dim ctAssayLevels As Short
                        'ctAssayLevels = tblAssayLevels.Rows.Count

                        ctAssayLevels = rowsAssays.Length


                        'ctAssayLevels doesn't always work
                        'E.G. AssayLevel=7 - Conc=2500, AssayLevel=7 - Conc=2600

                        'For Count3 = 0 To tblAssayLevels.Rows.Count - 1
                        '    ''''''console.writeline(tblAssayLevels.Rows(Count3).Item("ASSAYLEVEL"))
                        'Next

                        Dim int12 As Short = -1

                        int1 = InStr(strWRunId, " ", CompareMethod.Text)
                        If int1 = 0 Then
                            str2 = strWRunId
                        Else
                            str1 = Mid(strWRunId, 1, int1 - 1)
                            str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                        End If

                        If BOOLINCLUDEDATE Then
                            .Selection.Tables.Item(1).Cell(1, 1).Select()
                            If int1 = 0 Then
                            Else
                                .Selection.TypeText(str1)
                                .Selection.Tables.Item(1).Cell(2, 1).Select()
                            End If

                            .Selection.TypeText(str2)
                            .Selection.Tables.Item(1).Cell(3, 1).Select()
                            '.Selection.TypeText("(Analysis Date)")
                            '20180420 LEE:
                            .Selection.TypeText("(" & GetAnalysisDateLabel(intTableID) & ")")
                        Else
                            If int1 = 0 Then
                            Else
                                .Selection.Tables.Item(1).Cell(2, 1).Select()
                                .Selection.TypeText(str1)
                            End If

                            .Selection.Tables.Item(1).Cell(3, 1).Select()
                            .Selection.TypeText(str2)
                        End If

                        .Selection.Tables.Item(1).Cell(intHRow, 2).Select()

                        For Count2 = 1 To ctCalibrStds
                            If LboolNomConcParen Then
                                .Selection.TypeText(Text:="(" & CStr(strConcUnits) & ")")
                            Else
                                .Selection.TypeText(Text:=CStr(strConcUnits))
                            End If
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            If boolSTATSDIFFCOL Then
                                .Selection.TypeText(Text:=ReturnDiffLabel)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            End If
                        Next

                        '''''''wdd.visible = True
                        If boolSTATSREGR Then 'add regr col headings
                            For Count2 = 1 To intRP
                                Select Case Count2
                                    Case 1
                                        var1 = "A"
                                    Case 2
                                        var1 = "B"
                                    Case 3
                                        var1 = "C"
                                End Select
                                .Selection.TypeText(Text:=CStr(var1))

                                'superscript a
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:=" a")
                                .Selection.Font.Superscript = False
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                            Next

                            'do RSQ


                            'enter R2
                            var1 = "RSQ"
                            .Selection.TypeText(Text:=CStr(var1))
                            'superscript b
                            .Selection.Font.Superscript = True
                            .Selection.TypeText(Text:=" b")
                            .Selection.Font.Superscript = False
                            If boolSRegr Then
                            Else
                                'enter Regression
                                var1 = "Regr."
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                .Selection.TypeText(Text:=CStr(var1))
                                'superscript b
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:=" b")
                                .Selection.Font.Superscript = False

                            End If

                            If boolSWt Then
                            Else
                                'enter Weighting
                                var1 = "Wt"
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                .Selection.TypeText(Text:=CStr(var1))
                                'superscript b
                                .Selection.Font.Superscript = True
                                .Selection.TypeText(Text:=" b")
                                .Selection.Font.Superscript = False
                            End If

                            ''superscript 2
                            '.Selection.Font.Superscript = True
                            '.Selection.TypeText(Text:="2")
                            '.Selection.Font.Superscript = False
                            '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                            intLeg = intLeg + 1
                            intLegNNR = intLegNNR + 1
                            strA = Chr(intLeg + intLegStart)

                            Dim strL1 As String
                            For Count3 = 1 To intNumRegr

                                strRegressionType = arrRegrType(1, Count3)
                                'strWeighting = arrRegrType(2, Count3)

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
                                If boolSWt Then
                                    str1 = str2 & " where y is the peak area ratio of " & arrAnalytes(14, Count1) & " to Int. Std., x is the concentration of " & arrAnalytes(14, Count1) & ", and " & str3 & " are regression constants. Regression weighted " & strWeighting & "."
                                Else
                                    str1 = str2 & " where y is the peak area ratio of " & arrAnalytes(14, Count1) & " to Int. Std., x is the concentration of " & arrAnalytes(14, Count1) & ", and " & str3 & " are regression constants." ' Regression weighted " & strWeighting & "."
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

                            'add legend for RSQ
                            intLeg = intLeg + 1
                            intLegNNR = intLegNNR + 1

                            strA = Chr(intLeg + intLegStart)
                            arrLegend(1, intLeg) = strA
                            Dim strLA As String
                            strLA = "RSQ = R-Squared"
                            If boolSRegr = False Then
                                strLA = strLA & ", Regr. = Regression"
                            End If
                            If boolSWt = False Then
                                strLA = strLA & ", Wt = Weighting"
                            End If
                            arrLegend(2, intLeg) = strLA
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True
                            ctLegend = ctLegend + 1

                        End If

                        Dim intRow As Short
                        Dim intCol As Short

                        'enter concentration values
                        Count4 = intHRow + 1 '4 'Counter for row selection
                        Count5 = 0 'Counter for arr
                        intRow = Count4 ' 4
                        intCol = 0
                        .Selection.Tables.Item(1).Cell(Count4, 1).Select()

                        '''''''''''''''''wdd.visible = True

                        ''
                        ctP = 1

                        '***Start New
                        Dim numRunID As Int64
                        Dim numAssayLevel As Short
                        Dim rowsConc() As DataRow
                        Dim intID As Int64
                        Dim numRows As Short

                        intID = intAnalyteID

                        'herevis
                        ''''wdd.visible = True

                        For Count2 = 1 To intNumRuns  'number of RunIDs

                            numRunID = tblRunID.Rows(Count2 - 1).Item("RUNID") ' rows(Count2 - 1)

                            'enter RunID
                            If Count2 = ctP Then
                                'Select Case intTableID
                                '    Case 28
                                '        str1 = "Entering " & strTempInfo & " Final Extract Stability: Summary of Back Calculated Standard Concentrations Table for " & arrAnalytes(1, Count1) & "..."
                                '    Case Is <> 28
                                '        str1 = "Entering Back Calculated Calibration Standard Concentrations Table for " & arrAnalytes(1, Count1) & "..."
                                'End Select
                                str1 = strM
                                'frmH.lblProgress.Text = strM
                                'frmH.Refresh()
                                ctP = ctP + 5
                            End If
                            intRow = intRow + 1
                            intCol = 2 ' Count5 + 1
                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                            .Selection.TypeText(Text:=CStr(numRunID))
                            If BOOLINCLUDEDATE Then
                                .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()
                                str1 = GetDateFromRunID(NZ(numRunID, 0), LDateFormat, intGroup, idTR)
                                .Selection.TypeText("(" & str1 & ")")
                            End If
                            .Selection.Tables.Item(1).Cell(intRow, intCol).Select()

                            Dim boolEnterDiff As Boolean
                            boolEnterDiff = False

                            'Legend
                            '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag, 5=AnalyteFlagPercent
                            '6=Hi, 7=Lo, 8=varNom

                            'do each calibr level
                            Dim varConc
                            Dim varElim
                            Dim varDR 'Decision Reason
                            Dim varAFP
                            Dim varHi
                            Dim varLo
                            Dim maxReps As Short
                            Dim varcolorT
                            Dim boolNI As Boolean = False 'if standard level is included
                            Dim boolNV As Boolean = False 'if injection returns no value (e.g. NULL)

                            ''''''''wdd.visible = True

                            intCol = 1

                            maxReps = 0

                            'If numRunID = 31 Then
                            '    var1 = var1 'debug
                            'End If

                            int12 = 0

                            'check here
                            For Count3 = 0 To ctCalibrStds - 1 'ctAssayLevels

                                If numRunID = 21 And Count3 = 0 Then
                                    var1 = var1 'debug
                                End If

                                If boolSTATSDIFFCOL And Count3 > 0 Then
                                    intCol = intCol + 2
                                Else
                                    intCol = intCol + 1
                                End If

                                ''wdd.visible = True
                                boolNI = False

                                If Count3 > ctCalibrStds - 1 Then
                                    Exit For
                                End If

                                'numAssayLevel = tblNomConc.Rows(Count3).Item("LEVELNUMBER") ' tblAssayLevels.Rows(Count3).Item("ASSAYLEVEL")
                                numAssayLevel = tblNomConc.Rows(Count3).Item(strLevelNum) ' tblAssayLevels.Rows(Count3).Item("ASSAYLEVEL")

                                strF = "ANALYTEID = " & intID & " AND ASSAYLEVEL = " & numAssayLevel & " AND RUNID = " & numRunID
                                strS = "RUNSAMPLEORDERNUMBER ASC"
                                rowsConc = tblBCStdConcs1.Select(strF, strS)
                                numRows = rowsConc.Length
                                'varNom = tblNomConc.Rows(Count3).Item("Concentration")
                                varNom = tblNomConc.Rows(Count3).Item(strNomConc)
                                If numRows > maxReps Then
                                    maxReps = numRows
                                End If

                                If numRows = 0 Then

                                    .Selection.Tables.Item(1).Cell(intRow, intCol).Select()

                                    varcolorT = .Selection.Font.Color
                                    If boolRedBoldFont Then
                                        .Selection.Font.Bold = True
                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                    End If

                                    .Selection.TypeText(Text:="NI ")
                                    .Selection.Font.Bold = False
                                    .Selection.Font.Color = varcolorT ' Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                    varAFP = 15
                                    boolNI = True

                                Else

                                    intAssayID = NZ(rowsConc(0).Item("ASSAYID"), 0)
                                    Dim rowsAID() As DataRow
                                    If intRunNum = 0 Then 'NOT assigned samples
                                        'make sure this assayid corresponds to the nomconc

                                        'str1 = "CONCENTRATION = " & varNom & " AND ASSAYID = " & intAssayID & " AND ANALYTEID = " & intanalyteid
                                        ''str1 = "NOMCONC = " & varNom & " AND ASSAYLEVEL = " & numAssayLevel & " AND ANALYTEID = " & intanalyteid
                                        'Erase rowsAID
                                        ''cannot use tblBCStdsAssayID for assigned samples
                                        ''sometimes assaylevel gets duplicated, which is not reflected in tblBCStdsAssayID
                                        ''rowsAID = tblBCStdsAssayID.Select(str1)
                                        'rowsAID = tblBCStdsAssayID.Select(str1)

                                        'str1 = "NOMCONC = " & varNom & " AND ASSAYLEVEL = " & numAssayLevel & " AND ANALYTEID = " & intanalyteid
                                        str1 = "CONCENTRATION = " & varNom & " AND ASSAYID = " & intAssayID & " AND ANALYTEID = " & intAnalyteID
                                        Erase rowsAID
                                        'cannot use tblBCStdsAssayID for assigned samples
                                        'sometimes assaylevel gets duplicated, which is not reflected in tblBCStdsAssayID
                                        rowsAID = tblBCStdsAssayID.Select(str1)
                                        'rowsAID = tblBCStds1.Select(str1)
                                    Else 'ASSIGNED SAMPLES
                                        'make sure this assayid corresponds to the nomconc

                                        'str1 = "CONCENTRATION = " & varNom & " AND ASSAYID = " & intAssayID & " AND ANALYTEID = " & intanalyteid
                                        str1 = "NOMCONC = " & varNom & " AND ASSAYLEVEL = " & numAssayLevel & " AND ANALYTEID = " & intAnalyteID
                                        Erase rowsAID
                                        'cannot use tblBCStdsAssayID for assigned samples
                                        'sometimes assaylevel gets duplicated, which is not reflected in tblBCStdsAssayID
                                        'rowsAID = tblBCStdsAssayID.Select(str1)
                                        rowsAID = tblBCStds1.Select(str1)
                                    End If


                                    If rowsAID.Length = 0 Then
                                        '.Selection.Tables.item(1).cell(intRow, intCol).select()

                                        'If boolRedBoldFont Then
                                        '    .Selection.Font.Bold = True
                                        '    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                        'End If

                                        '.Selection.TypeText(Text:="NI ")
                                        'varAFP = 15
                                        boolNI = True
                                    Else

                                        '*****
                                        'find and record flag percent
                                        var1 = rowsConc(0).Item("ASSAYID") ' drows(Count3).Item("ASSAYID")
                                        int2 = tblBCStdsAssayID.Rows.Count 'debug
                                        If intRunNum = 0 Then
                                            str1 = "ASSAYID = " & var1 & " AND ANALYTEID = " & intAnalyteID & " AND LEVELNUMBER = " & numAssayLevel
                                        Else 'assigned samples
                                            'str1 = "ASSAYID = " & var1 & " AND ANALYTEID = " & intAnalyteID ' & " AND LEVELNUMBER = " & numAssayLevel
                                            str1 = "ASSAYID = " & var1 & " AND ANALYTEID = " & intAnalyteID & " AND LEVELNUMBER = " & numAssayLevel
                                        End If

                                        Dim strFP As String = str1

                                        str2 = "LEVELNUMBER ASC" ' "LEVELNUMBER ASC"
                                        rowsAssID = tblBCStdsAssayID.Select(str1, str2)
                                        Dim intRAIDrows As Short
                                        intRAIDrows = rowsAssID.Length 'debugging

                                        If boolEx Then 'if boolEx true, then is from assigned samples
                                            vU = NZ(rowsConc(0).Item("BOOLUSEGUWUACCCRIT"), 0)
                                            If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                                'do hi/lo later in code for assigned samples and GuWu Acc Crit

                                            Else

                                                '****

                                                If intRunNum = 0 Then

                                                    Try
                                                        var1 = rowsAssID(0).Item("ANALYTEFLAGPERCENT")
                                                        var3 = rowsAssID(0).Item("FLAGPERCENT")
                                                        var2 = NZ(var1, NZ(var3, 15)) 'debugging
                                                    Catch ex As Exception
                                                        var2 = 15 'System.DBNull.Value
                                                    End Try

                                                Else 'ASSIGNED SAMPLES

                                                    Try
                                                        boolRAID = False
                                                        'this is junk
                                                        'get from tblCalStdGroupsAll
                                                        'str1 = "RUNID = " & numRunID & " AND INTGROUP = " & intGroup & " AND LEVELNUMBER = " & numAssayLevel
                                                        str1 = "INTGROUP = " & intGroup & " AND LEVELNUMBER = " & numAssayLevel
                                                        Dim rowsAFP() As DataRow = tblCalStdGroupsAll.Select(str1)
                                                        If rowsAFP.Length = 0 Then
                                                            var2 = 15 'System.DBNull.Value
                                                        Else
                                                            'var2 = rowsAFP(0).Item("ANALYTEFLAGPERCENT")

                                                            '20170310 LEE: This is incorrect logic. Should be getting from individual analytical run
                                                            'in other words, rowsAssID

                                                            Try
                                                                var2 = GetFlagPercent(CDec(var1), intAnalyteID, numAssayLevel, CDec(varNom), numRunID)
                                                            Catch ex As Exception
                                                                Dim vvvv
                                                                vvvv = ex.Message
                                                            End Try

                                                            var2 = var2 'debug


                                                        End If

                                                    Catch ex As Exception
                                                        var2 = 15 'System.DBNull.Value
                                                    End Try
                                                End If

                                                '****

                                                varAFP = var2

                                                'Check Here
                                                Call SetHighAndLowCriteria(CDbl(varNom), CDbl(var2), CDbl(var2), varHi, varLo)

                                                v1 = var2
                                                v2 = var2

                                            End If
                                        Else

                                            '****

                                            If intRunNum = 0 Then

                                                Try
                                                    var1 = rowsAssID(0).Item("ANALYTEFLAGPERCENT")
                                                    var2 = NZ(rowsAssID(0).Item("ANALYTEFLAGPERCENT"), NZ(rowsAssID(0).Item("FLAGPERCENT"), 15)) 'debugging
                                                Catch ex As Exception
                                                    var2 = 15 'System.DBNull.Value
                                                End Try

                                            Else 'ASSIGNED SAMPLES

                                                Try
                                                    boolRAID = False

                                                    'this is junk
                                                    'get from tblCalStdGroupsAll
                                                    'str1 = "RUNID = " & numRunID & " AND INTGROUP = " & intGroup & " AND LEVELNUMBER = " & numAssayLevel
                                                    'RUNID in tblCalStdGroupsAll is an example only
                                                    str1 = "INTGROUP = " & intGroup & " AND LEVELNUMBER = " & numAssayLevel
                                                    Dim rowsAFP() As DataRow = tblCalStdGroupsAll.Select(str1)
                                                    If rowsAFP.Length = 0 Then
                                                        var2 = 15 'System.DBNull.Value
                                                    Else
                                                        'var2 = rowsAFP(0).Item("ANALYTEFLAGPERCENT")

                                                        '20170310 LEE: This is incorrect logic. Should be getting from individual analytical run
                                                        'in other words, rowsAssID

                                                        Try
                                                            var2 = GetFlagPercent(CDec(var1), intAnalyteID, numAssayLevel, CDec(varNom), numRunID)
                                                        Catch ex As Exception
                                                            Dim vvvv
                                                            vvvv = ex.Message
                                                        End Try

                                                    End If

                                                Catch ex As Exception
                                                    var2 = 15 'System.DBNull.Value
                                                End Try
                                            End If
                                            '****

                                            varAFP = var2

                                            'Check Here
                                            Call SetHighAndLowCriteria(CDbl(varNom), CDbl(var2), CDbl(var2), varHi, varLo)

                                            v1 = var2
                                            v2 = var2


                                        End If

                                        '*****
                                    End If

                                End If


                                Dim numConc As Single

                                ''wdd.visible = True


                                For Count6 = 0 To numRows - 1

                                    boolNoDiff = False
                                    boolOC = False

                                    If boolNI Then

                                        .Selection.Tables.Item(1).Cell(intRow + Count6, intCol).Select()
                                        boolOC = True

                                        varcolorT = .Selection.Font.Color
                                        If boolRedBoldFont Then
                                            .Selection.Font.Bold = True
                                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                        End If

                                        .Selection.TypeText(Text:="NI ")
                                        .Selection.Font.Color = varcolorT
                                        .Selection.Font.Bold = False

                                        boolNoDiff = True

                                    Else

                                        .Selection.Tables.Item(1).Cell(intRow + Count6, intCol).Select()

                                        'var1 = rowsConc(Count6).Item("ELIMINATEDFLAG") ' arrBCStdConcs(4, Count5)
                                        'var2 = rowsConc(Count6).Item("Concentration") 'arrBCStdConcs(2, Count5) 'debugging
                                        varElim = NZ(rowsConc(Count6).Item("ELIMINATEDFLAG"), "N") ' arrBCStdConcs(4, Count5)

                                        If boolNotAssignedSamples Then
                                            varDR = NZ(rowsConc(Count6).Item("DECISIONREASON"), "NA")
                                        Else
                                            'DECISIONREASON in not included in tblAssignedSamples
                                            'must retrieve from tblBCStdConcs
                                            'Also note that StudyDoc does not allow Excluding Calibr Stds in Assigned Samples interface
                                            'so don't have to evaluate BOOLEXCLSAMPLE
                                            str1 = "RUNID = " & numRunID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & numAssayLevel
                                            Dim rowsDR() As DataRow = tblBCStdConcs.Select(str1)
                                            var1 = rowsDR.Length 'debug
                                            If Count6 > rowsDR.Length - 1 Then
                                                varDR = "NA"
                                            Else
                                                varDR = NZ(rowsDR(Count6).Item("DECISIONREASON"), "NA")
                                            End If
                                            If StrComp(varDR, "NA", CompareMethod.Text) = 0 Then
                                                varDR = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical) and excluded from regression and summary statistics."
                                            Else
                                                varDR = varDR
                                            End If
                                        End If
                                        varConc = rowsConc(Count6).Item("Concentration") 'leave varconc as null if it is null
                                        var2 = NZ(varConc, "")
                                        If Len(var2) = 0 Then
                                            boolNV = True
                                            numConc = -1
                                        Else
                                            varConc = NZ(rowsConc(Count6).Item("Concentration"), 0) 'arrBCStdConcs(2, Count5) 'debugging
                                            If boolLUseSigFigs Then
                                                numConc = SigFigOrDec(varConc, LSigFig, False)
                                            Else
                                                numConc = RoundToDecimalRAFZ(varConc, LSigFig)
                                            End If
                                        End If


                                        ''wd.visible = True 'debugging

                                        Dim nrowRI As DataRow = tblRunIDNomConc.NewRow
                                        nrowRI.BeginEdit()
                                        nrowRI("NomConc") = varNom
                                        nrowRI("AssayID") = rowsConc(Count6).Item("ASSAYID")
                                        nrowRI.EndEdit()
                                        tblRunIDNomConc.Rows.Add(nrowRI)

                                        'If StrComp(varElim, "Y", vbTextCompare) = 0 Or IsDBNull(numConc) Then 
                                        If StrComp(varElim, "Y", vbTextCompare) = 0 And IsDBNull(varConc) = False Then

                                            boolNoDiff = True
                                            boolOC = True

                                            If boolCSSHOWREJVALUES Then

                                                'report value

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If IsDBNull(varConc) Then 'NO! Show NV!

                                                    boolNoDiff = True
                                                    boolNV = True
                                                    'report as NV No Value

                                                    str2 = Mid(varDR, Len(varDR), 1)
                                                    If StrComp(str2, ".", CompareMethod.Text) = 0 Then
                                                    Else
                                                        If boolSTATSREGR Then
                                                            varDR = "No Value: " & varDR ' & ". Value excluded from regression and summary statistics."
                                                        Else
                                                            varDR = "No Value: " & varDR ' & ". Value excluded from the regression and summary statistics."
                                                        End If
                                                    End If
                                                    str1 = varDR

                                                    intLeg = intLeg + 1
                                                    strA = Chr(intLeg + intLegStart)
                                                    ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)
                                                    If boolRedBoldFont Then
                                                        .Selection.Font.Bold = True
                                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                    End If

                                                    .Selection.TypeText(Text:="NV")

                                                    Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                                    .Selection.Font.Superscript = False
                                                    .Selection.Font.Bold = False

                                                ElseIf numConc = -1 Then 'this means there was no value recorded
                                                    .Selection.TypeText(Text:="NV")
                                                    boolNV = True
                                                ElseIf IsNumeric(numConc) Then
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(numConc, LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=CStr(Format(numConc, GetRegrDecStr(LSigFig))))
                                                    End If
                                                Else
                                                    If boolLUseSigFigs Then
                                                        .Selection.TypeText(Text:=CStr(DisplayNum(SigFigOrDec(0, LSigFig, False), LSigFig, False)))
                                                    Else
                                                        .Selection.TypeText(Text:=CStr(Format(0, GetRegrDecStr(LSigFig))))
                                                    End If

                                                End If

                                                If IsNumeric(numConc) And IsDBNull(varConc) = False Then
                                                    'herehere

                                                    If boolSTATSNR Then

                                                        'strA = strNR
                                                        'If intNR = 1 Then
                                                        '    strA = strNR
                                                        'Else
                                                        '    For Count4 = 1 To intNR
                                                        '        var1 = NZ(arrNR(1, Count4), -1)
                                                        '        If varAFP = var1 Then
                                                        '            'strA = arrNR(2, Count4)
                                                        '            'use (3 for this option
                                                        '            strA = arrNR(3, Count4)
                                                        '            Exit For
                                                        '        End If
                                                        '    Next
                                                        'End If

                                                        intLeg = intLeg + 1
                                                        strA = Chr(intLeg + intLegStart)

                                                        If boolEx Then 'if boolEx true, then is from assigned samples
                                                            vU = NZ(rowsConc(Count6).Item("BOOLUSEGUWUACCCRIT"), 0)
                                                            If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                                                v1 = NZ(rowsConc(Count6).Item("NUMMAXACCCRIT"), 0)
                                                                v2 = NZ(rowsConc(Count6).Item("NUMMINACCCRIT"), 0)

                                                                If boolSTATSREGR Then
                                                                    If v1 = v2 Then
                                                                        str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                                    Else
                                                                        str1 = "Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                                    End If
                                                                Else
                                                                    If v1 = v2 Then
                                                                        str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical) and excluded from the regression and summary statistics."
                                                                    Else
                                                                        str1 = "Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from the regression and summary statistics."
                                                                    End If
                                                                End If
                                                            Else

                                                                str2 = Mid(varDR, Len(varDR), 1)
                                                                If StrComp(str2, ".", CompareMethod.Text) = 0 Then
                                                                Else
                                                                    If StrComp(varDR, "NA", CompareMethod.Text) = 0 Then
                                                                        If boolSTATSREGR Then
                                                                            varDR = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                        Else
                                                                            varDR = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                        End If
                                                                    Else
                                                                        If boolSTATSREGR Then
                                                                            varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                        Else
                                                                            varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                        End If
                                                                    End If


                                                                End If
                                                                str1 = varDR

                                                            End If

                                                        Else
                                                            'str1 = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(varAFP, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                            'check for period
                                                            str2 = Mid(varDR, Len(varDR), 1)
                                                            If StrComp(str2, ".", CompareMethod.Text) = 0 Then
                                                                If boolSTATSREGR Then
                                                                    varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                Else
                                                                    varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                End If
                                                            Else
                                                                If StrComp(varDR, "NA", CompareMethod.Text) = 0 Then
                                                                    If boolSTATSREGR Then
                                                                        varDR = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                    Else
                                                                        varDR = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                    End If
                                                                Else
                                                                    If boolSTATSREGR Then
                                                                        varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                    Else
                                                                        varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                    End If
                                                                End If


                                                            End If
                                                            str1 = varDR
                                                        End If

                                                        'Add to Legend Array
                                                        ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                        If boolRedBoldFont Then
                                                            .Selection.Font.Bold = True
                                                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                        End If

                                                        Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                                    Else

                                                        intLeg = intLeg + 1
                                                        strA = Chr(intLeg + intLegStart)

                                                        If boolEx Then 'if boolEx true, then is from assigned samples
                                                            vU = NZ(rowsConc(Count6).Item("BOOLUSEGUWUACCCRIT"), 0)
                                                            If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                                                v1 = NZ(rowsConc(Count6).Item("NUMMAXACCCRIT"), 0)
                                                                v2 = NZ(rowsConc(Count6).Item("NUMMINACCCRIT"), 0)

                                                                If boolSTATSREGR Then
                                                                    If v1 = v2 Then
                                                                        str1 = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                                    Else
                                                                        str1 = "Not Reported: Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                                    End If
                                                                Else
                                                                    If v1 = v2 Then
                                                                        str1 = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical) and excluded from the regression and summary statistics."
                                                                    Else
                                                                        str1 = "Not Reported: Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from the regression and summary statistics."
                                                                    End If
                                                                End If

                                                            Else
                                                                str2 = Mid(varDR, Len(varDR), 1)
                                                                If StrComp(str2, ".", CompareMethod.Text) = 0 Then
                                                                Else
                                                                    If StrComp(varDR, "NA", CompareMethod.Text) = 0 Then
                                                                        If boolSTATSREGR Then
                                                                            varDR = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                        Else
                                                                            varDR = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                        End If
                                                                    Else
                                                                        If boolSTATSREGR Then
                                                                            varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                        Else
                                                                            varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                        End If
                                                                    End If


                                                                End If
                                                                str1 = varDR
                                                            End If

                                                            'str1 = "Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from regression and summary statistics."

                                                        Else
                                                            'str1 = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(varAFP, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                            'check for period
                                                            str2 = Mid(varDR, Len(varDR), 1)
                                                            If StrComp(str2, ".", CompareMethod.Text) = 0 Then
                                                            Else
                                                                If boolSTATSREGR Then
                                                                    varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from regression and summary statistics."
                                                                Else
                                                                    varDR = varDR & " (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical). Value excluded from the regression and summary statistics."
                                                                End If

                                                            End If
                                                            str1 = varDR
                                                        End If

                                                        'arrBCQCs(4, Count2)

                                                        'Add to Legend Array
                                                        ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                        'fontsize = .Selection.Font.Size

                                                        If boolRedBoldFont Then
                                                            .Selection.Font.Bold = True
                                                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                        End If

                                                        Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                                    End If
                                                Else

                                                End If


                                            Else

                                                'Here

                                                'boolNoDiff = True


                                                'strA = strNR
                                                'If intNR = 1 Then
                                                '    strA = strNR
                                                'Else
                                                '    For Count4 = 1 To intNR
                                                '        var1 = NZ(arrNR(1, Count4), -1)
                                                '        If varAFP = var1 Then
                                                '            strA = arrNR(2, Count4)
                                                '            Exit For
                                                '        End If
                                                '    Next
                                                'End If

                                                '.Selection.TypeText(Text:=" " & strA)

                                                intLeg = intLeg + 1
                                                strA = Chr(intLeg + intLegStart)


                                                '***record legend

                                                'strA = Chr(intLeg + intLegStart)
                                                'str1 = "Value outside of acceptance criteria (" & RoundToDecimal(varAFP, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                'str1 = "Not reported - Standard is outside of acceptance criteria (" & RoundToDecimal(arrBCStdConcs(5, Count5), 0) & "% theoretical) and excluded from regression and summary statistics."

                                                If boolEx Then 'if boolEx true, then is from assigned samples
                                                    vU = NZ(rowsConc(Count6).Item("BOOLUSEGUWUACCCRIT"), 0)
                                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                                        v1 = NZ(rowsConc(Count6).Item("NUMMAXACCCRIT"), 0)
                                                        v2 = NZ(rowsConc(Count6).Item("NUMMINACCCRIT"), 0)
                                                    End If

                                                    If boolSTATSREGR Then
                                                        If v1 = v2 Then
                                                            str1 = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                        Else
                                                            str1 = "Not Reported: Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                        End If
                                                    Else
                                                        If v1 = v2 Then
                                                            str1 = "Not Reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & "% theoretical) and excluded from the regression and summary statistics."
                                                        Else
                                                            str1 = "Not Reported: Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from the regression and summary statistics."
                                                        End If
                                                    End If


                                                    'str1 = "Not reported - Standard is outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & "% theoretical) and excluded from regression and summary statistics."

                                                Else
                                                    'str1 = "Not reported: Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(varAFP, 0) & "% theoretical) and excluded from regression and summary statistics."
                                                    str2 = Mid(varDR, Len(varDR), 1)
                                                    If StrComp(str2, ".", CompareMethod.Text) = 0 Then
                                                    Else
                                                        If StrComp(varDR, "NA", CompareMethod.Text) = 0 Then
                                                            If boolSTATSREGR Then
                                                                varDR = "Not Reported: Value excluded from regression and summary statistics."
                                                            Else
                                                                varDR = "Not Reported: Value excluded from the regression and summary statistics."
                                                            End If
                                                        Else
                                                            If boolSTATSREGR Then
                                                                varDR = "Not Reported: " & varDR & ". Value excluded from regression and summary statistics."
                                                            Else
                                                                varDR = "Not Reported: " & varDR & ". Value excluded from the regression and summary statistics."
                                                            End If
                                                        End If


                                                    End If
                                                    str1 = varDR
                                                End If

                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                ''add NR to arrLegend(1
                                                'Dim vL1, vL2
                                                'vL1 = arrLegend(1, intLeg)
                                                'vL2 = "NR " & vL1
                                                'arrLegend(1, intLeg) = vL2

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                .Selection.TypeText(Text:="NR")

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                                '.Selection.TypeText(Text:=" NR")
                                                .Selection.Font.Superscript = False
                                                .Selection.Font.Bold = False

                                            End If

                                            boolEnterDiff = True 'FALSE

                                        ElseIf IsDBNull(varConc) Then

                                            boolNoDiff = True
                                            boolNV = True
                                            boolOC = True

                                            'report as NV No Value

                                            'intLeg = intLeg + 1
                                            'strA = Chr(intLeg + intLegStart)

                                            'str1 = "No Value: Standard not acquired in this injection"

                                            str2 = Mid(varDR, Len(varDR), 1)
                                            If StrComp(str2, ".", CompareMethod.Text) = 0 Then
                                            Else
                                                If StrComp(varDR, "NA", CompareMethod.Text) = 0 Then
                                                    If boolSTATSREGR Then
                                                        varDR = "No Value" ': " & varDR ' & ". Value excluded from regression and summary statistics."
                                                    Else
                                                        varDR = "No Value" ': " & varDR ' & ". Value excluded from the regression and summary statistics."
                                                    End If
                                                Else
                                                    If boolSTATSREGR Then
                                                        varDR = "No Value: " & varDR ' & ". Value excluded from regression and summary statistics."
                                                    Else
                                                        varDR = "No Value: " & varDR ' & ". Value excluded from the regression and summary statistics."
                                                    End If
                                                End If

                                            End If
                                            str1 = varDR

                                            intLeg = intLeg + 1
                                            strA = Chr(intLeg + intLegStart)
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)
                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If

                                            .Selection.TypeText(Text:="NV")

                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                            .Selection.Font.Superscript = False
                                            .Selection.Font.Bold = False

                                            'arrBCQCs(4, Count2)

                                            'fontsize = .Selection.Font.Size

                                        Else
                                            'determine if value is outside acceptance criteria
                                            If boolLUseSigFigs Then
                                                var1 = CStr(DisplayNum(numConc, LSigFig, False)) ' CStr(numConc) ' DisplayNum(CDbl(numConc), LSigFig, False) 'conc value
                                            Else
                                                var1 = CStr(Format(numConc, GetRegrDecStr(LSigFig))) ' CStr(numConc) ' DisplayNum(CDbl(numConc), LSigFig, False) 'conc value
                                            End If

                                            var2 = varHi 'arrBCStdConcs(6, Count5) 'Hello
                                            var3 = varLo 'arrBCStdConcs(7, Count5) 'Lo
                                            boolEnterDiff = True
                                            boolNoDiff = False

                                            If boolEx Then 'if boolEx true, then is from assigned samples
                                                vU = NZ(rowsConc(Count6).Item("BOOLUSEGUWUACCCRIT"), 0)
                                                If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                                    v1 = NZ(rowsConc(Count6).Item("NUMMAXACCCRIT"), 0)
                                                    v2 = NZ(rowsConc(Count6).Item("NUMMINACCCRIT"), 0)
                                                    Call SetHighAndLowCriteria(varNom, v1, v2, var2, var3)
                                                End If
                                            End If

                                            '.Selection.TypeText(Text:=CStr(var1))
                                            ''don't do this for calibr stds!!!
                                            'WHY NOT??

                                            '20160506 LEE: Begin using +/-AccCrit

                                            'Check Here
                                            'If CDec(var1) > var2 Or CDec(var1) < var3 Then 'flag
                                            If OutsideAccCrit(numConc, varNom, v1, v2, NZ(vU, 0)) Then

                                                intLeg = intLeg + 1
                                                intLegNNR = intLegNNR + 1

                                                Try
                                                    'strA = Chr(intLegNNR + intLegStart)
                                                    strA = Chr(intLeg + intLegStart)
                                                Catch ex As Exception

                                                    strA = "z"
                                                End Try

                                                If boolEx Then 'if boolEx true, then is from assigned samples
                                                    vU = NZ(rowsConc(Count6).Item("BOOLUSEGUWUACCCRIT"), 0)
                                                    If gAllowGuWuAccCrit And LAllowGuWuAccCrit And vU = -1 Then
                                                        If boolSTATSREGR Then
                                                            If v1 = v2 Then
                                                                str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & " % theoretical) but included in regression and summary statistics."
                                                            Else
                                                                str1 = "Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & " % theoretical) but included in regression and summary statistics."
                                                            End If
                                                        Else
                                                            If v1 = v2 Then
                                                                str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimalRAFZ(v1, 0) & " % theoretical) but included in summary statistics."
                                                            Else
                                                                str1 = "Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1, 0) & "/-" & RoundToDecimalRAFZ(v2, 0) & " % theoretical) but included in summary statistics."
                                                            End If
                                                        End If


                                                        'str1 = "Value outside of acceptance criteria (+" & RoundToDecimalRAFZ(v1) & "/-" & RoundToDecimalRAFZ(v2) & "% theoretical) but included in regression and summary statistics."

                                                    Else
                                                        If boolSTATSREGR Then
                                                            str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(varAFP) & "% theoretical) but included in regression and summary statistics."
                                                        Else
                                                            str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(varAFP) & "% theoretical) but included in summary statistics."
                                                        End If

                                                    End If
                                                Else

                                                    If boolSTATSREGR Then
                                                        str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(varAFP) & "% theoretical) but included in regression and summary statistics."
                                                    Else
                                                        str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(varAFP) & "% theoretical) but included in summary statistics."
                                                    End If

                                                End If

                                                '
                                                'arrBCQCs(4, Count2)
                                                'search for str1 in arrLegend

                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                'fontsize = .Selection.Font.Size

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                .Selection.TypeText(Text:=CStr(var1))

                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                                .Selection.Font.Superscript = False
                                                .Selection.Font.Bold = False

                                            Else

                                                .Selection.TypeText(Text:=CStr(var1))

                                            End If


                                        End If

                                        If boolSTATSDIFFCOL And boolNoDiff = False Then

                                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                            If boolEnterDiff Then
                                                'var3 = Format(((numConc / varNom) - 1) * 100, strQCDec)
                                                var3 = Format(RoundToDecimal(((numConc / varNom) - 1) * 100, intQCDec), strQCDec)
                                                If boolTHEORETICAL Then
                                                    'var3 = CalcREPercent(var2, varNom, intQCDec)
                                                    var3 = CalcREPercent(numConc, varNom, intQCDec)
                                                    numTheor = 100 + CDec(var3)

                                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3 + 1, "Accuracy", numTheor, CSng(numRunID), Count1, strDo, v1, v2, boolOC)

                                                Else
                                                    'var3 = Format(RoundToDecimal(CalcREPercent(var2, varNom, intQCDec), intQCDec), strQCDec)
                                                    var3 = Format(RoundToDecimal(CalcREPercent(numConc, varNom, intQCDec), intQCDec), strQCDec)

                                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count3 + 1, "Accuracy", var3, CSng(numRunID), Count1, strDo, v1, v2, boolOC)

                                                End If
                                            Else

                                                If boolQCNA Then
                                                    var3 = "NA"
                                                Else
                                                    var3 = ""
                                                End If

                                            End If
                                            .Selection.TypeText(Text:=CStr(var3))
                                            '.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        End If

                                    End If


                                Next

                                '''''''''wdd.visible = True

                                If Count3 = ctCalibrStds - 1 Then 'ctAssayLevels
                                    If boolSTATSREGR Then 'add regression information
                                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                        For Count7 = 1 To intRP
                                            var1 = arrRegCon(Count7 + 1, intRPIncr)
                                            var2 = NZ(var1, "NA")
                                            If IsNumeric(var2) Then
                                                'Dim Count8 As Short
                                                'str2 = ""
                                                'For Count8 = 1 To LRegrSigFigs - 1
                                                '    str2 = str2 & "0"
                                                'Next
                                                'str2 = "0." & str2 & "E+0"
                                                'str1 = Format((SigFigOrDec(var2, LRegrSigFigs, False, False)), str2)
                                                str2 = var2
                                                str1 = Format(var2, str2)
                                                .Selection.TypeText(Text:=str1)
                                                '.Selection.TypeText(Text:=CStr(SigFigOrDec(var2, LRegrSigFigs, False, False)))

                                            Else
                                                .Selection.TypeText(Text:=CStr(var2))
                                            End If
                                            '.Selection.TypeText(Text:=CStr(SigFigOrDec(var2, LRegrSigFigs, False, False)))
                                            .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                                        Next
                                        '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                                        var1 = arrRegCon(intRP + 2, intRPIncr)
                                        var2 = var1 'CStr(SigFigOrDec(var1, LR2SigFigs, False, False))'ALREADY FORMATTED
                                        .Selection.TypeText(Text:=CStr(NZ(var2, "NA")))

                                        '''''''wdd.visible = True

                                        .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                                        If boolSRegr Then
                                        Else
                                            'enter regression
                                            var2 = arrRunID(intRPIncr) 'returns runid
                                            var1 = GetRegrRegCon(CInt(var2), intAnalyteID)
                                            .Selection.TypeText(Text:=var1)
                                            .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                                        End If

                                        If boolSWt Then
                                        Else
                                            'enter weighting
                                            'var1 = "1/X^2"
                                            var2 = arrRunID(intRPIncr) 'returns runid
                                            var1 = GetWtRegCon(CInt(var2), intAnalyteID)
                                            .Selection.TypeText(Text:=var1)
                                            '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                                        End If


                                        intRPIncr = intRPIncr + 1

                                    End If

                                End If

                            Next

                            intRow = intRow + maxReps 'numRows

                        Next

                        intRow = intRow + 1
                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                        'begin doing normal statistics
                        'rather than moving down, must use cell numbers from this point
                        'because if the stats section is near a page break, VB has a spaz
                        int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)

                        Dim intCell As Short
                        Dim intCols As Short
                        Dim intHome As Short
                        intCell = intRow 'Count4 '.Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber)
                        intCols = ctCalibrStds + 1
                        int1 = 0

                        '''''''''''''''wdd.visible = True

                        If boolCSREPORTACCVALUES Then
                        Else
                            If intExp > 0 Then
                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                '.Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                                .Selection.TypeText(Text:="Summary Statistics Excluding 'Not Reported' Values")
                                '.Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                Try
                                    .Selection.Cells.Merge()
                                Catch ex As Exception

                                End Try
                                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                End With
                                '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                                int1 = int1 + 1
                                .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                                '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                            End If

                        End If
                        '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                        '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        'Mean, SD, %CV, %Bias
                        intCell = intCell + int1
                        intHome = intCell ' + int1
                        int1 = -1

                        '''''''''''''''wdd.visible = True


                        If boolSTATSMEAN Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            .Selection.TypeText(Text:="Mean")
                        End If
                        If boolSTATSSD Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                            .Selection.TypeText(Text:="S.D.")
                        End If
                        If boolSTATSCV Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                            .Selection.TypeText(Text:=ReturnPrecLabel())
                        End If
                        If boolSTATSBIAS And boolSTATSMEAN Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                            .Selection.TypeText(Text:="%Bias")
                        End If
                        If boolSTATSDIFF And boolSTATSMEAN Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                            .Selection.TypeText(Text:="%Diff")
                        End If
                        If BOOLSTATSRE And boolSTATSMEAN Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                            .Selection.TypeText(Text:="%RE")
                        End If
                        If boolTHEORETICAL And boolSTATSMEAN Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                            .Selection.TypeText(Text:="%Theoretical")
                        End If

                        If boolSTATSN Then
                            int1 = int1 + 1
                            If intCell + int1 > .Selection.Tables.Item(1).Rows.Count Then
                                .Selection.InsertRowsBelow(1)
                            End If
                            .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()
                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                            .Selection.TypeText(Text:="n")
                        End If

                        .Selection.Tables.Item(1).Cell(intHome, 2).Select()
                        int1 = 0

                        '''''''''''''''wdd.visible = True


                        '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)
                        '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)

                        'ReDim arrBCStdActual(inttemprows * 2)

                        'The next portion is to get num Mean

                        ReDim arrBCStdActual(inttemprows)
                        ctP = 1

                        int12 = 0
                        'For Count3 = 1 To ctCalibrStds * int11 Step int11
                        For Count3 = 0 To ctCalibrStds - 1 'ctAssayLevels

                            int12 = int12 + 1
                            If Count3 = ctP Then
                                strM = "Entering " & strTName & " Statistics For " & arrAnalytes(1, Count1) & " for Level " & Count3 & " of " & ctAssayLevels & " calibration stds..."
                                strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                                frmH.lblProgress.Text = strM
                                frmH.Refresh()
                                ctP = ctP + 5
                            End If

                            If Count3 > ctCalibrStds - 1 Then
                                Exit For
                            End If

                            'numAssayLevel = tblNomConc.Rows(Count3).Item("LEVELNUMBER")
                            numAssayLevel = tblNomConc.Rows(Count3).Item(strLevelNum) 'tblNomConc comes from tblCalStdGroupsAcc
                            'numNomConc = tblNomConc.Rows(Count3).Item("Concentration")
                            numNomConc = tblNomConc.Rows(Count3).Item(strNomConc)

                            'int1 = 0
                            'herehere
                            If intRunNum = 0 Then 'NOT assigned samples
                                'If boolIncludePSAE Or boolEx Then
                                '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12 & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS = 3"
                                'Else
                                '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3"
                                'End If

                                'If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                                '    str1 = "ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & numAssayLevel & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                                'Else
                                '    str1 = "ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & numAssayLevel & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                                'End If

                                If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                                    str1 = "ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & numAssayLevel & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS = 3"
                                Else
                                    str1 = "ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & numAssayLevel & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3"
                                End If

                                'have to add ASSAYID's for some occasions

                                For Count2 = 0 To rowsAssays.Length - 1
                                    var1 = rowsAssays(Count2).Item("AssayID")
                                    If Count2 = 0 Then
                                        str2 = " AND (ASSAYID = " & var1
                                    Else
                                        str2 = str2 & " OR ASSAYID = " & var1
                                    End If
                                Next
                                If rowsAssays.Length = 0 Then
                                Else
                                    str2 = str2 & ")"
                                    str1 = str1 & str2
                                End If

                                'Dim rowsNI() As DataRow
                                'strF = "NomConc = " & numNomConc
                                'rowsNI = tblRunIDNomConc.Select(strF)
                                'Dim Count1a As Short
                                'For Count1a = 0 To rowsNI.Length - 1
                                '    var1 = rowsNI(Count1a).Item("AssayID")
                                '    If Count1a = 0 Then
                                '        str2 = " AND (ASSAYID = " & var1
                                '    Else
                                '        str2 = str2 & " OR ASSAYID = " & var1
                                '    End If
                                'Next
                                'If rowsNI.Length = 0 Then
                                'Else
                                '    str2 = str2 & ")"
                                '    str1 = str1 & str2
                                'End If

                            Else

                                'str1 = "ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & numAssayLevel & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                                str1 = "ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & numAssayLevel
                                'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12
                                'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND CONCENTRATION = " & arrBCStds(2, Count3)
                                str1 = str1 & " AND ("
                                int1 = rows.Length
                                str2 = ""
                                For Count5 = 0 To int1 - 1
                                    var1 = rows(Count5)
                                    If Count5 = int1 - 1 Then
                                        str2 = str2 & "RUNID = " & var1 & ")"
                                    Else
                                        str2 = str2 & "RUNID = " & var1 & " Or "
                                    End If
                                Next
                                str1 = str1 & str2

                            End If

                            Erase drows

                            drows = tblBCStdConcs1.Select(str1)
                            intN = 0
                            int2 = drows.Length
                            ReDim arrBCStdActual(int2)
                            'var2 = ""
                            For Count5 = 0 To int2 - 1
                                'num1 = NZ(drows(Count3).Item("CONCENTRATION"), 0)
                                'num1 = SigFigOrDec(CDec(num1), LSigFig, True)
                                'arrBCStdConcs(2, Count3 + 1) = num1
                                'arrBCStdConcs(3, Count3 + 1) = drows(Count3).Item("RUNID")
                                'arrBCStdConcs(4, Count3 + 1) = drows(Count3).Item("ELIMINATEDFLAG")
                                var1 = NZ(drows(Count5).Item("ELIMINATEDFLAG"), "N") 'arrBCStdConcs(4, Count5)
                                If StrComp(var1, "Y", vbTextCompare) = 0 Or IsDBNull(drows(Count5).Item("CONCENTRATION")) Then 'exclude value
                                Else
                                    intN = intN + 1
                                    num1 = NZ(drows(Count5).Item("CONCENTRATION"), 0)
                                    num2 = NZ(drows(Count5).Item("ALIQUOTFACTOR"), 0)
                                    num3 = CDbl(num1 / num2)
                                    num1 = SigFigOrDec(num3, LSigFig, False)

                                    arrBCStdActual(intN) = num1
                                    'var2 = var2 & num1 & ";"
                                    'var7 = arrBCStdConcs(2, Count5)
                                End If
                            Next

                            '''''''''''''console.writeline(var2)
                            'determine Sum
                            numSum = 0

                            If boolLUseSigFigs Then
                                numMean = SigFigOrDec(Mean(intN, arrBCStdActual), LSigFig, False)
                                numSD = SigFigOrDec(StdDev(intN, arrBCStdActual), LSigFig, False)
                            Else
                                numMean = RoundToDecimalRAFZ(Mean(intN, arrBCStdActual), LSigFig)
                                numSD = RoundToDecimalRAFZ(StdDev(intN, arrBCStdActual), LSigFig)
                            End If


                            int1 = -1

                            ''''''''wdd.visible = True


                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Mean", numMean, CSng(numRunID), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSMEAN Then
                                int1 = int1 + 1
                                If int11 = 1 Or Count3 = 0 Then
                                    .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                Else
                                    .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                End If

                                Try
                                    'record mean
                                    If intN = 0 Then
                                        .Selection.TypeText(Text:="NA")
                                    Else
                                        If boolLUseSigFigs Then
                                            .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                        Else
                                            .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                        End If
                                    End If


                                Catch ex As Exception

                                End Try
                            End If


                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "SD", numSD, CSng(numRunID), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSSD Then
                                Try
                                    int1 = int1 + 1
                                    If int11 = 1 Or Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                    Else
                                        .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                    End If
                                    '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                    'record SD
                                    If intN < gSDMax Or intN = 0 Then
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
                                If intN < gSDMax Then
                                Else
                                    numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Precision", numPrec, CSng(numRunID), Count1, strDo, 0, 0, False)
                                End If

                            Catch ex As Exception

                            End Try
                            If boolSTATSCV Then
                                Try
                                    int1 = int1 + 1
                                    If int11 = 1 Or Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                    Else
                                        .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                    End If
                                    'record %CV
                                    If intN < gSDMax Or intN = 0 Then
                                        .Selection.TypeText("NA")
                                    Else

                                        .Selection.TypeText(Format(numPrec, strQCDec))

                                    End If

                                Catch ex As Exception

                                End Try
                            End If


                            If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, numNomConc, intQCDec)
                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", numBias, CSng(numRunID), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else

                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, False)

                            End If

                            If boolSTATSBIAS And boolSTATSMEAN Then
                                Try
                                    int1 = int1 + 1
                                    If int11 = 1 Or Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                    Else
                                        .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                    End If

                                    If intN = 0 Then
                                        .Selection.TypeText("NA")
                                    Else
                                        .Selection.TypeText(Format(numBias, strQCDec))
                                    End If


                                Catch ex As Exception

                                End Try
                            End If


                            If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numTheor = CalcREPercent(numMean, numNomConc, intQCDec)
                                    numTheor = 100 + CDec(numTheor)
                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", numTheor, CSng(numRunID), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numTheor = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, False)
                                numTheor = 100 + CDec(numTheor)
                            End If

                            If boolTHEORETICAL And boolSTATSMEAN Then
                                Try
                                    int1 = int1 + 1
                                    If int11 = 1 Or Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                    Else
                                        .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                    End If

                                    If intN = 0 Then
                                        .Selection.TypeText("NA")
                                    Else
                                        .Selection.TypeText(Format(numTheor, strQCDec))
                                    End If

                                Catch ex As Exception

                                End Try
                            End If


                            If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, numNomConc, intQCDec)

                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "Accuracy", numBias, CSng(numRunID), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, False)
                            End If

                            If boolSTATSDIFF And boolSTATSMEAN Then
                                Try
                                    int1 = int1 + 1
                                    If int11 = 1 Or Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                    Else
                                        .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                    End If

                                    If intN = 0 Then
                                        .Selection.TypeText("NA")
                                    Else
                                        .Selection.TypeText(Format(numBias, strQCDec))
                                    End If

                                Catch ex As Exception

                                End Try
                            End If


                            If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, numNomConc, intQCDec)

                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", numBias, CSng(numRunID), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, False)
                            End If

                            If BOOLSTATSRE And boolSTATSMEAN Then
                                Try
                                    int1 = int1 + 1
                                    If int11 = 1 Or Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                    Else
                                        .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                    End If

                                    If intN = 0 Then
                                        .Selection.TypeText("NA")
                                    Else
                                        .Selection.TypeText(Format(numBias, strQCDec))
                                    End If

                                Catch ex As Exception

                                End Try
                            End If


                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12 + 1, "n", intN, CSng(numRunID), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSN Then
                                Try
                                    int1 = int1 + 1
                                    If int11 = 1 Or Count3 = 0 Then
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 2).Select()
                                    Else
                                        .Selection.Tables.Item(1).Cell(intHome + int1, ((Count3 + 2) * int11) - 2).Select()
                                    End If
                                    '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                    'record n
                                    .Selection.TypeText(Text:=CStr(intN))

                                Catch ex As Exception

                                End Try
                            End If

                            intCell = intHome + int1
                            'int1 = int1 + 1
                            'int1 = 0

                            '''''''wdd.visible = True

                            If Count3 >= ctCalibrStds * int11 Then
                            Else
                                '.selection.Tables.item(1).cell(intHome, Count3 + int11 + 1).select()
                                int1 = 0

                                'If Count3 + int11 + 1 > .selection.tables.item(1).columns.count Then
                                'Else
                                '    .selection.Tables.item(1).cell(intHome, Count3 + int11 + 2).select()
                                'End If

                                '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)
                                '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                            End If


                        Next

                        'intCell = .Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber)
                        'intCell = intHome + int1
                        intCols = ctCalibrStds + 1

                        ''''''''''wdd.visible = True

                        'boolSTATSREGR = False
                        If boolSTATSREGR Then 'add stuff for regression

                            Dim intN1 As Short
                            'legend: ReDim arrRegCon(intRP + 2, intRegrCt)

                            For Count3 = 1 To intRP + 1
                                intN1 = 0
                                num2 = 0
                                For Count5 = 1 To intRegrCt

                                    'intN1 = intN1 + 1
                                    'num1 = arrRegr(Count3, Count5)
                                    'var1 = NZ(arrRegr(Count3, Count5), "NA")
                                    var2 = arrRegCon(Count3 + 1, Count5)
                                    var1 = NZ(var2, "NA")
                                    If IsNumeric(var1) Then
                                        intN1 = intN1 + 1
                                        num1 = CSng(var1)
                                        If intN1 > UBound(arrBCStdActual, 1) Then
                                            ReDim Preserve arrBCStdActual(intN1)
                                        End If
                                        arrBCStdActual(intN1) = num1
                                    Else
                                        'arrBCStdActual(intN1) = var1
                                    End If

                                Next

                                '''''''''''''console.writeline(var2)
                                'determine Sum
                                numSum = 0
                                If Count3 = intRP + 1 Then 'this is R2
                                    If boolLUseSigFigsRegr Then
                                        numMean = SigFigOrDec(Mean(intN1, arrBCStdActual), LR2SigFigs, False)
                                    Else
                                        numMean = RoundToDecimalRAFZ(Mean(intN1, arrBCStdActual), LR2SigFigs)
                                    End If
                                    'numMean = SigFigOrDec(Mean(intN1, arrBCStdActual), LR2SigFigs, False)
                                Else
                                    numMean = SigFigOrDec(Mean(intN1, arrBCStdActual), LRegrSigFigs, False)
                                End If
                                'numMean = SigFigOrDec(Mean(intN1, arrBCStdActual), LRegrSigFigs, False)
                                'numSD = SigFigOrDec(StdDev(intN1, arrBCStdActual), LSigFig, False)
                                If boolLUseSigFigsRegr Then
                                    numSD = SigFigOrDec(StdDev(intN1, arrBCStdActual), LRegrSigFigs, False)
                                Else
                                    numSD = RoundToDecimalRAFZ(StdDev(intN1, arrBCStdActual), LRegrSigFigs)
                                End If


                                int1 = -1
                                If boolSTATSMEAN Then
                                    '''''''''''wdd.visible = True
                                    Try
                                        'record mean
                                        int1 = int1 + 1
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1 + (ctCalibrStds * int11)).Select()
                                        '.selection.Tables.item(1).cell(intHome + int1, Count3 + 2).select()

                                        'Dim Count8 As Short
                                        'str2 = ""
                                        'For Count8 = 1 To LRegrSigFigs - 1
                                        '    str2 = str2 & "0"
                                        'Next
                                        'str2 = "0." & str2 & "E+0"
                                        str2 = GetScNot(LRegrSigFigs)
                                        'str1 = Format((SigFigOrDec(var2, LRegrSigFigs, False, False)), str2)
                                        var2 = numMean ' DisplayNum(numMean, LRegrSigFigs, False)
                                        If IsNumeric(var2) And Count3 <> intRP + 1 Then
                                            str1 = Format(CDec(var2), str2)
                                        ElseIf IsNumeric(var2) And Count3 = intRP + 1 Then 'this is for R2
                                            'str1 = CStr(var2)
                                            str1 = Format(CDec(var2), GetScNot(LR2SigFigs))
                                        Else
                                            str1 = CStr(var2)
                                        End If

                                        If intN1 = 0 Then
                                            .Selection.TypeText("NA")
                                        Else
                                            .Selection.TypeText(Text:=str1)
                                        End If


                                    Catch ex As Exception

                                    End Try
                                End If
                                If boolSTATSSD Then
                                    Try
                                        int1 = int1 + 1
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1 + (ctCalibrStds * int11)).Select()
                                        '.selection.Tables.item(1).cell(intHome + int1, Count3 + 2).select()
                                        '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                        'record SD
                                        If intN1 < gSDMax Then
                                            .Selection.TypeText("NA")
                                        Else

                                            'Dim Count8 As Short
                                            'str2 = ""
                                            'For Count8 = 1 To LSigFig - 1
                                            '    str2 = str2 & "0"
                                            'Next
                                            'str2 = "0." & str2 & "E+0"
                                            str2 = GetScNot(LRegrSigFigs)
                                            'str1 = Format((SigFigOrDec(var2, LRegrSigFigs, False, False)), str2)
                                            var2 = numSD ' DisplayNum(numSD, LRegrSigFigs, False)
                                            If IsNumeric(var2) And Count3 <> intRP + 1 Then
                                                str1 = Format(CDec(var2), str2)
                                            ElseIf IsNumeric(var2) And Count3 = intRP + 1 Then 'this is for R2
                                                'str1 = CStr(var2)
                                                str1 = Format(CDec(var2), GetScNot(LR2SigFigs))
                                            Else
                                                str1 = CStr(var2)
                                            End If

                                            If intN1 = 0 Then
                                                .Selection.TypeText("NA")
                                            Else
                                                .Selection.TypeText(Text:=str1)
                                            End If
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolSTATSCV Then
                                    Try
                                        int1 = int1 + 1
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1 + (ctCalibrStds * int11)).Select()
                                        '.selection.Tables.item(1).cell(intHome + int1, Count3 + 2).select()
                                        '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                        'record %CV
                                        If intN1 < gSDMax Or intN1 = 0 Then
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
                                        int1 = int1 + 1

                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolTHEORETICAL And boolSTATSMEAN Then
                                    Try
                                        int1 = int1 + 1

                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolSTATSDIFF And boolSTATSMEAN Then
                                    Try
                                        int1 = int1 + 1

                                    Catch ex As Exception

                                    End Try
                                End If

                                If BOOLSTATSRE And boolSTATSMEAN Then
                                    Try
                                        int1 = int1 + 1

                                    Catch ex As Exception

                                    End Try
                                End If

                                If boolSTATSN Then
                                    Try
                                        int1 = int1 + 1
                                        .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1 + (ctCalibrStds * int11)).Select()
                                        '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                        'record n
                                        .Selection.TypeText(Text:=CStr(intN1))

                                    Catch ex As Exception

                                    End Try
                                End If
                            Next

                            .Selection.Tables.Item(1).Cell(intCell, 1).Select()
                            int1 = 0

                        End If


                        '''''''''''''''wdd.visible = True

                        'int1 = 0

                        If boolCSREPORTACCVALUES Then
                            'border bottom of thellos table
                            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        Else
                            If intExp > 0 Then
                            Else
                                'border bottom of thellos table
                                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                                .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            End If
                        End If

                        'begin doing statistics that includes outliers
                        'rather than moving down, must use cell numbers from this point
                        'because if the stats section is near a page break, VB has a spaz
                        If boolCSREPORTACCVALUES Then
                        Else
                            If intExp > 0 Then
                                intCell = intCell + 2 '.Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber)
                                int1 = Count4
                                int1 = 0
                                'int1 = .Selection.Information(Microsoft.Office.Interop.Word.wdinformation.wdStartOfRangeRowNumber)

                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)

                                .Selection.Tables.Item(1).Cell(intCell + int1, 1).Select()

                                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)

                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                '.Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                                .Selection.TypeText(Text:="Summary Statistics Including 'Not Reported' Values")
                                '.Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                Try
                                    .Selection.Cells.Merge()
                                Catch ex As Exception

                                End Try
                                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                End With

                                int1 = int1 + 1
                                intHome = intCell + int1
                                int1 = -1
                                .Selection.Tables.Item(1).Cell(intHome + int1, 1).Select()

                                '.Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                                '.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)

                                '''''''''''''''wdd.visible = True

                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)

                                Call typeStatsLabels(wd, int1, intHome, 1, False)

                                .Selection.Tables.Item(1).Cell(intHome, 2).Select()

                                'ReDim arrBCStdActual(inttemprows * 2)
                                ReDim arrBCStdActual(inttemprows)

                                ctP = 1
                                int12 = 0
                                For Count3 = 1 To ctCalibrStds * int11 Step int11
                                    int12 = int12 + 1
                                    If Count3 = ctP Then
                                        strM = "Entering " & strTName & " Statistics For " & arrAnalytes(1, Count1) & " for Level " & int12 & " of " & ctCalibrStds & " calibration stds..."
                                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                                        frmH.lblProgress.Text = strM
                                        frmH.Refresh()
                                        ctP = ctP + 5
                                    End If
                                    'int1 = 0
                                    'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & Count3
                                    If intRunNum = 0 Then 'NOT assigned samples
                                        'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & Count3
                                        'If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                                        '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12 & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                                        'Else
                                        '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                                        'End If

                                        If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                                            str1 = "ASSAYLEVEL = " & int12 & " AND RUNTYPEID > 0 And RUNANALYTEREGRESSIONSTATUS = 3 AND INTGROUP = " & intGroup
                                        Else
                                            str1 = "ASSAYLEVEL = " & int12 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND INTGROUP = " & intGroup
                                        End If

                                    Else

                                        'If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                                        '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12 & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                                        'Else
                                        '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                                        'End If

                                        If boolIncludePSAE Or boolEx Then 'if boolEx true, then is from assigned samples
                                            str1 = "ASSAYLEVEL = " & int12 & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS = 3 AND INTGROUP = " & intGroup
                                        Else
                                            str1 = "ASSAYLEVEL = " & int12 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND INTGROUP = " & intGroup
                                        End If
                                        'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & int12
                                        'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND CONCENTRATION = " & arrBCStds(2, Count3)
                                        str1 = str1 & " AND ("
                                        int1 = rows.Length
                                        str2 = ""
                                        For Count5 = 0 To int1 - 1
                                            var1 = rows(Count5)
                                            If Count5 = int1 - 1 Then
                                                str2 = str2 & "RUNID = " & var1 & ")"
                                            Else
                                                str2 = str2 & "RUNID = " & var1 & " OR "
                                            End If
                                        Next
                                        str1 = str1 & str2
                                    End If
                                    Erase drows
                                    drows = tblBCStdConcs1.Select(str1)
                                    intN = 0
                                    int2 = drows.Length
                                    ReDim arrBCStdActual(int2)
                                    For Count5 = 0 To int2 - 1
                                        'num1 = NZ(drows(Count3).Item("CONCENTRATION"), 0)
                                        'num1 = SigFigOrDec(CDec(num1), LSigFig, True)
                                        'arrBCStdConcs(2, Count3 + 1) = num1
                                        'arrBCStdConcs(3, Count3 + 1) = drows(Count3).Item("RUNID")
                                        'arrBCStdConcs(4, Count3 + 1) = drows(Count3).Item("ELIMINATEDFLAG")
                                        var1 = NZ(drows(Count5).Item("ELIMINATEDFLAG"), "N") 'arrBCStdConcs(4, Count5)
                                        'If StrComp(var1, "Y", vbTextCompare) = 0 Then 'exclude value
                                        'Else
                                        intN = intN + 1
                                        'num1 = NZ(drows(Count5).Item("CONCENTRATION"), 0)
                                        'num1 = SigFigOrDec(CDec(num1), LSigFig, True)
                                        num1 = NZ(drows(Count5).Item("CONCENTRATION"), 0)
                                        num2 = NZ(drows(Count5).Item("ALIQUOTFACTOR"), 1)
                                        num3 = CDbl(num1 / num2)
                                        num1 = SigFigOrDec(num3, LSigFig, False)

                                        arrBCStdActual(intN) = num1
                                        'var7 = arrBCStdConcs(2, Count5)
                                        'End If
                                    Next
                                    'determine Sum
                                    numSum = 0
                                    If boolLUseSigFigs Then
                                        numMean = SigFigOrDec(Mean(intN, arrBCStdActual), LSigFig, False)
                                        numSD = SigFigOrDec(StdDev(intN, arrBCStdActual), LSigFig, False)
                                    Else
                                        numMean = RoundToDecimalRAFZ(Mean(intN, arrBCStdActual), LSigFig)
                                        numSD = RoundToDecimalRAFZ(StdDev(intN, arrBCStdActual), LSigFig)
                                    End If
                                    int1 = -1

                                    If boolSTATSMEAN Then
                                        Try
                                            'record mean
                                            int1 = int1 + 1
                                            If boolLUseSigFigs Then
                                                .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                            Else
                                                .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                            End If


                                        Catch ex As Exception

                                        End Try
                                    End If
                                    If boolSTATSSD Then
                                        Try
                                            int1 = int1 + 1
                                            .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1).Select()
                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            'record SD
                                            If intN < gSDMax Then
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

                                    varNom = CDec(arrBCStds(2, int12))
                                    If boolSTATSCV Then
                                        Try

                                            int1 = int1 + 1
                                            .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1).Select()
                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            'record %CV
                                            If intN < gSDMax Then
                                                .Selection.TypeText("NA")
                                            Else
                                                numPrec = CalcCVPercent(numSD, numMean, intQCDec)


                                                .Selection.TypeText(Format(numPrec, strQCDec))
                                            End If

                                        Catch ex As Exception

                                        End Try
                                    End If


                                    If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numBias = CalcREPercent(numMean, numNomConc, intQCDec)
                                        Catch ex As Exception

                                        End Try
                                    Else

                                        'get numbias from average of %Bias columns
                                        numBias = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, True)
                                    End If

                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        Try
                                            int1 = int1 + 1
                                            .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1).Select()
                                            .Selection.TypeText(Format(numBias, strQCDec))

                                        Catch ex As Exception

                                        End Try
                                    End If


                                    If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numTheor = CalcREPercent(numMean, numNomConc, intQCDec)
                                            numTheor = 100 + CDec(numTheor)
                                        Catch ex As Exception

                                        End Try
                                    Else

                                        'get numbias from average of %Bias columns
                                        numTheor = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, True)
                                        numTheor = 100 + CDec(numTheor)
                                    End If

                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        Try
                                            int1 = int1 + 1
                                            .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1).Select()
                                            .Selection.TypeText(Format(numTheor, strQCDec))
                                        Catch ex As Exception

                                        End Try
                                    End If


                                    If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numBias = CalcREPercent(numMean, numNomConc, intQCDec)
                                        Catch ex As Exception

                                        End Try
                                    Else

                                        'get numbias from average of %Bias columns
                                        numBias = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, True)
                                    End If

                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        Try
                                            int1 = int1 + 1
                                            .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1).Select()
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                        Catch ex As Exception

                                        End Try
                                    End If


                                    If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                        Try
                                            numBias = CalcREPercent(numMean, numNomConc, intQCDec)
                                        Catch ex As Exception

                                        End Try
                                    Else

                                        'get numbias from average of %Bias columns
                                        numBias = GetBiasFromDiffCol(idTR, numNomConc, int12, 0, True)
                                    End If
                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        Try
                                            int1 = int1 + 1
                                            .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1).Select()
                                            .Selection.TypeText(Format(numBias, strQCDec))

                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolSTATSN Then
                                        Try
                                            int1 = int1 + 1
                                            .Selection.Tables.Item(1).Cell(intHome + int1, Count3 + 1).Select()
                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            'record n
                                            .Selection.TypeText(Text:=CStr(intN))



                                        Catch ex As Exception

                                        End Try
                                    End If

                                    intCell = intHome + int1

                                    If Count3 >= ctCalibrStds * int11 Then
                                    Else
                                        If Count3 + int11 + 1 > .Selection.Tables.Item(1).Columns.Count Then
                                        Else
                                            .Selection.Tables.Item(1).Cell(intHome, Count3 + int11 + 1).Select()
                                        End If

                                        int1 = 0

                                        '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)
                                        '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                                    End If
                                Next

                                .Selection.Tables.Item(1).Cell(intCell, 1).Select()
                                int1 = 0

                            End If


                        End If
                        'end stats section that includes outliers

                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try



                    'border bottom of the table
                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    ''''''wdd.visible = True

                    str1 = "Final formatting of " & strTName & " for " & arrAnalytes(1, Count1) & "..."
                    strM = str1
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    'check for multiple strNA's
                    var1 = intNRPos
                    var2 = intNRUsed

                    If intNR <> 1 And intNRUsed = 1 Then
                        arrLegend(1, intNRPos) = "NR"

                        'now do a search/replace
                        .Selection.Tables.Item(1).Select()
                        wd.Selection.Find.ClearFormatting()
                        wd.Selection.Find.Replacement.ClearFormatting()
                        Dim CountRR As Short
                        For CountRR = 1 To 2
                            Select Case CountRR
                                Case 1
                                    str1 = "NRa"
                                Case 2
                                    str1 = "NRb"
                            End Select
                            With wd.Selection.Find
                                .Text = str1
                                .Replacement.Text = "NR"
                                .Forward = True
                                '.Wrap = wdFindAsk
                                .Format = False
                                .MatchCase = False
                                .MatchWholeWord = True
                                '.MatchWildcards = False
                                '.MatchSoundsLike = False
                                '.MatchAllWordForms = False
                            End With
                            wd.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                        Next

                    End If

                    'enter table number
                    str1 = "Summary of " & strAnalC & " Back-Calculated Calibration Standard Concentrations Table"

                    'autofit table
                    Call AutoFitTable(wd, False)


                    'remove unused rows
                    Call RemoveRows(wd, 1)

                    '***
                    If gNumMatrix = 1 Then
                        strA = strAnalC
                    Else
                        strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                    End If
                    'No. Now just send strAnal
                    strA = strAnal
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 5, strA, strTempInfo, intTableID, intGroup, idTR)
                    'Note: strTName is byRef and will return Table, number, caption, label

                    '***

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

                    'intLeg = intLeg + 1
                    'arrLegend(1, intLeg) = " NR"
                    'arrLegend(2, intLeg) = "Not reported - Standard is outside acceptance criteria and excluded from regression and summary statistics"
                    'arrLegend(3, intLeg) = False
                    'ctLegend = intLeg

                    'intLeg = intLeg + 1
                    'arrLegend(1, intLeg) = " NRA"
                    'arrLegend(2, intLeg) = "Not reported - Standard is outside acceptance criteria (> 20%) and excluded from regression and summary statistics"
                    'arrLegend(3, intLeg) = False
                    'ctLegend = intLeg

                    'intLeg = intLeg + 1
                    'arrLegend(1, intLeg) = " NRB"
                    'arrLegend(2, intLeg) = "Not reported - Standard is outside acceptance criteria (> 15%) and excluded from regression and summary statistics"
                    'arrLegend(3, intLeg) = False
                    'ctLegend = intLeg

                    intLeg = intLeg + 1
                    arrLegend(1, intLeg) = "NI "
                    arrLegend(2, intLeg) = "Not Included: Standard level not included in this calibration run"
                    arrLegend(3, intLeg) = False
                    ctLegend = intLeg

                    '20160208 LEE: Need 'No Value' for AbbVie data
                    'If data record is null, then NV = true
                    ''20160223 LEE: No, NV incorporated earlier
                    'intLeg = intLeg + 1
                    'arrLegend(1, intLeg) = "NV "
                    'arrLegend(2, intLeg) = "No Value: Standard not acquired in this injection"
                    'arrLegend(3, intLeg) = False
                    'ctLegend = intLeg


                    intLeg = intLeg + 1
                    arrLegend(1, intLeg) = "NA"
                    arrLegend(2, intLeg) = "Not Applicable"
                    arrLegend(3, intLeg) = False
                    ctLegend = intLeg

                    ReDim Preserve arrLegend(4, ctLegend)

                    ''evaluate NR's
                    ''DOESN'T WORK. 'NRB' is already typed in the table
                    'Dim intNNNR As Short = 0
                    'Dim intNNNS As Short = 0

                    'For Count2 = 1 To intLeg
                    '    var1 = arrLegend(1, Count2)
                    '    If InStr(1, var1, "NR", CompareMethod.Text) > 0 Then
                    '        intNNNR = intNNNR + 1
                    '        intNNNS = Count2
                    '    End If
                    'Next
                    'If intNNNR = 1 Then
                    '    arrLegend(1, intNNNS) = "NR"
                    'End If

                    str1 = frmH.lblProgress.Text
                    '
                    ''''''wdd.visible = True



                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean, ByVal boolCarefulSplit As Boolean, ByVal boolFirstAnova As Boolean, ByVal intTableID As Int64)

                    'autofit table
                    Dim boolTT As Boolean
                    If BOOLINCLUDEDATE Or boolSTATSREGR Then
                        boolTT = True
                    Else
                        boolTT = False
                    End If
                    Call AutoFitTable(wd, boolTT)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    Call SplitTable(wd, 4, ctLegend, arrLegend, str1, False, ctLegend, False, True, False, intTableID)

                    'autofit table
                    Call AutoFitTable(wd, False)

                    ''remove unused rows
                    'Call RemoveRows(wd, 1)

                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)

                    var1 = "a" 'debugging

                End If
end1:

next1:

skip3:
            Next Count1

        End With


    End Sub


    Sub SRSummaryOfIQCCR_UseGroups_4(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64)


        Dim boolOC As Boolean = False 'bool if eliminated
        Dim numNomConc As Decimal
        Dim BACStudy As String
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
        Dim arrTemp(2, 50)
        Dim num1 As Single
        Dim num2 As Single
        Dim num3 As Single
        Dim num4 As Single
        Dim arrBCStdActual(1, 1)
        Dim arrBCQCActual(1)

        Dim ctLegend As Short
        Dim lng1 As Integer
        Dim lng2 As Integer
        Dim boolPortrait As Boolean
        Dim intLastAnal As Short
        Dim arrOrder(1, 1)
        Dim ctCols 'number of columns in a table
        Dim strSub1 As String
        Dim strSub2 As String
        Dim pos1 As Short
        Dim pos2 As Short
        Dim numSum As Object
        Dim numMean As Object
        Dim numSD As Object
        Dim rsF As New ADODB.Recordset
        Dim maxRep As Short
        Dim ctQCLegend As Short = 0
        'ctQCLegend = 0

        Dim var10
        Dim varConc

        Dim arrQCLegend(4, 20)
        Dim ctDilLeg As Short
        Dim ctQCAI As Short
        'Dim frmp As New frmprogress_01
        Dim dvDo As System.Data.DataView
        Dim intDo As Short
        Dim strDo As String
        Dim bool As Boolean
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim dv As System.Data.DataView
        Dim drows() As DataRow
        Dim drowsF() As DataRow
        Dim arrBCQCs(1, 1) '1=LevelNumber, 2=Concentration
        '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel
        Dim arrBCQCConcs(1, 1)
        Dim int10 As Short
        Dim intF As Short
        Dim Count10 As Short
        Dim ctP As Short
        Dim boolGo As Boolean
        Dim boolPro As Boolean
        Dim ctQCs As Short
        Dim arrQCAI(3, 500) '1=arrBCQCs number, 2=AssayID,3=NomConcentration
        Dim DilQCFactor(20)
        Dim inttemprows As Short
        Dim ctAnalyticalRuns As Short
        Dim strTName As String
        Dim var100
        Dim intLegStart As Short
        Dim strA As String
        Dim intLeg As Short
        Dim strB As String
        Dim arrFP(4, 20) 'flag percent array
        '1=max, 2=min, 3=hi, 4=lo
        Dim fontsize
        Dim numLevels As Short
        Dim strTempInfo As String
        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short
        Dim intCS As Short
        Dim intCE As Short
        Dim strS As String
        Dim intLevel As Short
        Dim hi As Double
        Dim lo As Double

        Dim ctTbl As Short = 0

        Dim v1, v2, vU

        Dim tblZ As New System.Data.DataTable

        Dim numPrec As Single
        Dim numBias As Single
        Dim numTheor As Single
        Dim varNom
        Dim strConcUnits As String
        Dim strDecReason As String

        Dim charFCID As String
        Dim strF As String
        strF = "ID_TBLREPORTTABLE = " & idTR
        Dim rowsTR() As DataRow = tblReportTable.Select(strF)
        var1 = rowsTR(0).Item("CHARFCID")
        charFCID = NZ(var1, "NA")

        ''wdd.visible = True

        Dim fonts
        fontsize = wd.Selection.Font.Size
        fonts = wd.Selection.Font.Size

        'Group variables
        Dim intGroup As Short
        Dim intAnalyteIndex As Int64
        Dim intAnalyteID As Int64
        Dim strAnalayteFlag As String
        Dim strMatrix As String
        Dim strAnal As String
        Dim strAnalC As String
        Dim tblAG As DataTable = tblAnalyteGroups 'tblAnalyteGroups has all analytes, not just accepted
        Dim intSpecies As Short
        Dim intCR As Short
        Dim strTNameO As String 'original Table Name
        Dim strFAssayID As String

        Dim strFFF As String

        'Note that tblBCQCConcs has NomConc and QCLABEL and AnalyteDescription columns added

        With wd

            'dvDo = frmH.dgvReportTableConfiguration.DataSource
            'strTName = "Summary of Interpolated QC Std Conc"
            'intDo = FindRowDVByCol(strTName, dvDo, "Table")


            Dim intTableID As Short
            intTableID = 4

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

            For Count1 = 1 To ctAnalytes

                Dim arrLegend(4, 20)

                strTName = strTNameO

                ctLegend = 0

                gnumAnal = Count1

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


                '* Check if table is to be generated for this Analyte
                intGroup = tblAG.Rows(Count1 - 1).Item("INTGROUP") ' arrAnalytes(15, Count1)
                'intAnalyteIndex = arrAnalytes(3, Count1)'NOTE:  may be more than one analyteindex per group!!!
                intAnalyteID = tblAG.Rows(Count1 - 1).Item("ANALYTEID") ' arrAnalytes(2, Count1)
                strMatrix = tblAG.Rows(Count1 - 1).Item("MATRIX")
                strAnal = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION")
                strAnalC = tblAG.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION_C")

                strDo = strAnalC ' arrAnalytes(1, Count1) 'record column name (Analyte Description)

                gstrAnal = strAnal

                If UseAnalyte(CStr(strDo)) Then
                Else
                    GoTo next1
                End If

                ''get an example AssayID from tblcalstdgroupassayid

                'str1 = "INTGROUP = " & intGroup
                ''ensure item is selected in anal run summary
                'strFFF = GetARSRuns(tblCalStdGroupAssayIDsAcc, intAnalyteID, "")

                'If Len(strFFF) = 0 Then
                'Else
                '    strFFF = "(" & strFFF & ")"
                '    str1 = str1 & " AND " & strFFF
                'End If

                'If BOOLINCLUDEDATE Then
                '    str2 = "RUNDATE ASC, RUNID ASC"
                'Else
                '    str2 = "RUNID ASC"
                'End If

                'Dim rowsAllRuns() As DataRow = tblCalStdGroupAssayIDsAcc.Select(str1, str2)
                'If rowsAllRuns.Length = 0 Then
                '    GoTo end1
                'End If
                'Dim intAssayID As Int64
                'intAssayID = rowsAllRuns(0).Item("ASSAYID")
                'ctAnalyticalRuns = rowsAllRuns.Length

                bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column

                Dim strM As String
                Dim strM1 As String
                If bool Then 'continue

                    intTCur = intTCur + 1

                    ctTbl = ctTbl + 1

                    strM = "Creating " & strTName & " For " & strAnalC & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()

                    If tblZ.Columns.Contains("Nomconc") Then
                        tblZ.Clear()
                    Else
                        'build tblz to record stats info
                        tblZ.Columns.Add("NomConc", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("Conc", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("ALIQUOTFACTOR", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("ELIMINATEDFLAG", Type.GetType("System.String"))
                        tblZ.Columns.Add("BOOLOUTLIER", Type.GetType("System.Boolean"))
                        tblZ.Columns.Add("HI", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("LO", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("RunID", Type.GetType("System.Int16"))
                        tblZ.Columns.Add("FlagPercent", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("numRep", Type.GetType("System.Int16"))
                        tblZ.Columns.Add("DECISIONREASON", Type.GetType("System.String"))
                        tblZ.Columns.Add("v1", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("v2", Type.GetType("System.Decimal"))
                        tblZ.Columns.Add("QCLABEL", Type.GetType("System.String"))
                        tblZ.Columns.Add("ASSAYLEVEL", Type.GetType("System.Decimal"))

                    End If

                    '******

                    'get an example AssayID from tblcalstdgroupassayid

                    str1 = "INTGROUP = " & intGroup
                    'ensure item is selected in anal run summary
                    strFFF = GetARSRuns(tblCalStdGroupAssayIDsAcc, intAnalyteID, "", False)

                    If Len(strFFF) = 0 Then
                    Else
                        strFFF = "(" & strFFF & ")"
                        str1 = str1 & " AND " & strFFF
                    End If

                    If BOOLINCLUDEDATE Then
                        str2 = "RUNDATE ASC, RUNID ASC"
                    Else
                        str2 = "RUNID ASC"
                    End If

                    Dim rowsAllRuns() As DataRow = tblCalStdGroupAssayIDsAcc.Select(str1, str2)
                    If rowsAllRuns.Length = 0 Then
                        GoTo end1
                    End If
                    Dim intAssayID As Int64
                    intAssayID = rowsAllRuns(0).Item("ASSAYID")
                    ctAnalyticalRuns = rowsAllRuns.Length

                    '******


                    ''20160221 LEE: Don't need anymore
                    'int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                    'str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                    'If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                    'Else
                    '    strConcUnits = str1
                    'End If

                    'page setup according to configuration
                    Dim strOrientation As String
                    strOrientation = dvDo.Item(intDo).Item("CHARPAGEORIENTATION")
                    'insert page break
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Call InsertPageBreak(wd)
                    Call PageSetup(wd, strOrientation) 'L=Landscape, P=Portrait

                    'ReDim arrBCQCs(10, 50) 
                    '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel
                    strM = "Creating " & strTName & " For " & strAnalC & "..."
                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    strM1 = strM
                    frmH.lblProgress.Text = strM
                    frmH.Refresh()


                    Dim tblLevelCrit As New DataTable
                    Dim col1 As New DataColumn
                    col1.ColumnName = "NomConc"
                    col1.DataType = System.Type.GetType("System.Decimal")
                    tblLevelCrit.Columns.Add(col1)
                    Dim col2 As New DataColumn
                    col2.ColumnName = "Crit"
                    col2.DataType = System.Type.GetType("System.Decimal")
                    tblLevelCrit.Columns.Add(col2)


                    'find number of QC levels

                    ' ''20160221 LEE: Don't need anymore
                    'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1)
                    ''strS = "CONCENTRATION ASC" 'bad if there is a dilution sample
                    'strS = "LEVELNUMBER ASC"
                    ''strS = "CONCENTRATION ASC"
                    'drows = tblBCQCs.Select(str1, strS)
                    'int1 = drows.Length

                    'must go through a convoluted process because users can really screw things up in Watson
                    'first get all accepted analytical runs for this group
                    '  - got it here: rowsAllRuns
                    'now loop through to get data only for these runs

                    strFAssayID = GetASSAYIDFilterIDCT(intGroup, False, True, intTableID)

                    strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID
                    Dim rowsQCAllRuns() As DataRow = tblBCQCsAssayID.Select(strF)
                    'convert this to table
                    Dim tblQCAllRuns As DataTable = rowsQCAllRuns.CopyToDataTable
                    ''now get unique concentrations
                    'Dim dvQCAllRuns As DataView = New DataView(tblQCAllRuns, "", "ASSAYID ASC, CONCENTRATION ASC", DataViewRowState.CurrentRows)
                    '20160318 LEE: DILUTIONFACTOR has been added to tblBCQCsAssayID
                    'Dim dvQCAllRuns As DataView = New DataView(tblQCAllRuns, "", "ASSAYID ASC, CONCENTRATION ASC, DILUTIONFACTOR DESC", DataViewRowState.CurrentRows) 'DilutionFactor DESC because Watson DilF is inverse
                    Dim dvQCAllRuns As DataView = New DataView(tblQCAllRuns, "", "CONCENTRATION ASC, DILUTIONFACTOR DESC, ID ASC", DataViewRowState.CurrentRows) 'DilutionFactor DESC because Watson DilF is inverse
                    'Dim tblNomConc As DataTable = dvQCAllRuns.ToTable("a", True, "LEVELNUMBER", "CONCENTRATION", "ID") 'BIG Assumption: these all runs have the same levelnumber for corresponding nomconc
                    '20160222 LEE: Can't do LEVELNUMBER here. Sometimes user have same Conc, but different LEVELNUMBER in different Assays. See AbbVie A1041392_C2 
                    '20160318 LEE: DILUTIONFACTOR has been added to tblBCQCsAssayID, a more logical filter
                    'Dim tblNomConc As DataTable = dvQCAllRuns.ToTable("a", True, "CONCENTRATION", "ID") 'BIG Assumption: these all runs have the same levelnumber for corresponding nomconc
                    Dim tblNomConc As DataTable = dvQCAllRuns.ToTable("a", True, "CONCENTRATION", "ID", "DILUTIONFACTOR") 'BIG Assumption: these all runs have the same levelnumber for corresponding nomconc

                    ''debug
                    ''console.writeline("Start: " & strFAssayID)
                    'var1 = ""
                    'For Count3 = 0 To tblNomConc.Columns.Count - 1
                    '    var2 = tblNomConc.Columns(Count3).ColumnName
                    '    var1 = var1 & ChrW(9) & var2
                    'Next
                    ''console.writeline(var1)

                    'For Count2 = 0 To tblNomConc.Rows.Count - 1
                    '    var1 = ""
                    '    For Count3 = 0 To tblNomConc.Columns.Count - 1
                    '        var2 = tblNomConc.Rows(Count2).Item(Count3)
                    '        var1 = var1 & ChrW(9) & var2
                    '    Next
                    '    'console.writeline(var1)
                    'Next
                    ''console.writeline("End: " & strFAssayID)

                    'record concentration units
                    strConcUnits = rowsAllRuns(0).Item("CONCENTRATIONUNITS")

                    int1 = tblNomConc.Rows.Count

                    ReDim Preserve arrFP(4, int1)

                    Erase arrBCQCs
                    ReDim arrBCQCs(10, int1)
                    '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel

                    Dim drowsAA() As DataRow

                    Dim vAnalyteIndex
                    Dim vMasterAssayID
                    Dim vAnalyteID
                    Dim varAC

                    '20171124 LEE: OK, here's the scoop for complicated studies
                    'tblNomConc can have NULL for DILUTIONFACTOR, which can mess up sorts
                    'Resolution: Convert NULL to 1 at this point
                    For Count2 = 0 To tblNomConc.Rows.Count - 1
                        var1 = NZ(tblNomConc.Rows(Count2).Item("DILUTIONFACTOR"), 1)
                        tblNomConc.Rows(Count2).BeginEdit()
                        tblNomConc.Rows(Count2).Item("DILUTIONFACTOR") = var1
                        tblNomConc.Rows(Count2).EndEdit()
                    Next

                    'convert tblnomconc to drows to conserve code
                    strF = "CONCENTRATION > 0"
                    'drows = tblNomConc.Select(strF, "CONCENTRATION ASC")
                    '20171124 LEE: Account for multiple DilnF
                    'drows = tblNomConc.Select(strF, "CONCENTRATION ASC, DILUTIONFACTOR DESC, ID ASC")
                    'drows = tblNomConc.Select(strF, "CONCENTRATION ASC, DILUTIONFACTOR DESC")

                    '20171128 LEE:
                    'Goofy QC Scenario: LI00016 - ACHN172001
                    'has a Carryover QC that gets stuck in between two dilution QCs
                    'need to switch sort
                    drows = tblNomConc.Select(strF, "DILUTIONFACTOR DESC, CONCENTRATION ASC, ID ASC")

                    '20160318 LEE:
                    'tblNomConc may contain DilnFactors that aren't actually used in the study
                    'there is not a robust way to create an appropriate query at the Watson level (tblBCQCsAssayID) that excludes non-used dilnfactors
                    'tried to add ANALYTICALRUNSAMPLE to query, but this requires a join between ASSAYLEVEL and LEVELNUMBER, which is dangerous
                    'instead, apply a filter to the next code that determines the different QC levels
                    Count2 = 0
                    For Count3 = 0 To drows.Length - 1

                        'before recording, ensure that qc level was ever used
                        'var3 = drows(Count3).Item("LevelNumber")

                        'Can only check nomconc
                        var4 = drows(Count3).Item("Concentration")
                        var5 = intAnalyteID ' drows(Count3).Item("ANALYTEID")
                        Try
                            var7 = drows(Count3).Item("DILUTIONFACTOR") 'debug
                            var6 = CDec(NZ(drows(Count3).Item("DILUTIONFACTOR"), 1))
                        Catch ex As Exception
                            var6 = 1

                        End Try

                        'vAnalyteIndex = drows(Count3).Item("AnalyteIndex")
                        'vMasterAssayID = drows(Count3).Item("MasterAssayID")
                        vAnalyteID = var5

                        ''str1 = "ASSAYLEVEL = " & var3 & " AND ANALYTEINDEX = " & var2 & " and MASTERASSAYID = " & var1
                        ''str1 = "ASSAYLEVEL = " & var3 & " AND ANALYTEINDEX = " & var2 & " and MASTERASSAYID = " & var1 & " AND NOMCONC = " & var4
                        ''str1 = "ASSAYLEVEL = " & var3 & " AND NOMCONC = " & var4 & " AND ANALYTEID = " & var5
                        'str1 = "NOMCONC = " & var4 & " AND ANALYTEID = " & var5

                        'sometimes a level is present that isn't used.
                        'must filter this out
                        If var6 = 1 Then
                            strF = strFAssayID & " AND NOMCONC = " & var4 & " AND ANALYTEID = " & var5
                        Else
                            strF = strFAssayID & " AND NOMCONC = " & var4 & " AND ANALYTEID = " & var5 & " AND ALIQUOTFACTOR = " & var6
                        End If

                        Erase drowsAA
                        drowsAA = tblBCQCConcs.Select(strF)
                        int2 = drowsAA.Length
                        If int2 = 0 Then
                        Else
                            Count2 = Count2 + 1
                            var1 = NZ(drows(Count3).Item("CONCENTRATION"), 0)
                            numNomConc = var1

                            'var2 = drows(Count3).Item("LevelNumber")
                            'can't get from drows anymore
                            'get from drowsAA.ASSAYLEVEL
                            'var2 = drowsAA(Count3).Item("ASSAYLEVEL")
                            var2 = drowsAA(0).Item("ASSAYLEVEL")

                            'arrBCQCs(9, int1)
                            '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 
                            '8=ASSAYID, 9=Aliquotfactor

                            arrBCQCs(1, Count2) = var2 ' 
                            arrBCQCs(2, Count2) = var1
                            var3 = drows(Count3).Item("ID")
                            arrBCQCs(3, Count2) = drows(Count3).Item("ID") 'might be able to filter on this later
                            var4 = drowsAA(0).Item("QCLABEL")
                            arrBCQCs(10, Count2) = drowsAA(0).Item("QCLABEL") 'same as ID

                            '****

                            ''skip all this. Will be done later
                            ''determine hi and lo (nom*flagpercent)
                            'Note: Actually, for QCs, this logic is true
                            Dim rows10() As DataRow

                            'strF = "CONCENTRATION = " & var1 & " AND ANALYTEID = " & vAnalyteID & " AND MASTERASSAYID = " & vMasterAssayID & " AND ANALYTEINDEX = " & vAnalyteIndex & " AND CONCENTRATION = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3" ' & " AND RUNID = " & var10
                            'rows10 = tblQCRunIDs.Select(strF, "RUNID ASC, LEVELNUMBER ASC")

                            'strF = "CONCENTRATION = " & var1 & " AND ANALYTEID = " & vAnalyteID & " AND CONCENTRATION = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3 AND LEVELNUMBER = " & var2

                            strF = "ANALYTEID = " & vAnalyteID & " AND CONCENTRATION = " & var1 & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            'Note: Filter doesn't need strFAssayID because tblQCAllRuns has already been filtered for this
                            'rows10 = tblQCAllRuns.Select(strF, "RUNID ASC, LEVELNUMBER ASC")

                            'if nomConc < 1, then the query return 0 records
                            'must do something different
                            strF = "ANALYTEID = " & vAnalyteID & " AND RUNANALYTEREGRESSIONSTATUS = 3"
                            'Note: Filter doesn't need strFAssayID because tblQCAllRuns has already been filtered for this
                            rows10 = tblQCAllRuns.Select(strF, "RUNID ASC, LEVELNUMBER ASC")

                            Dim intRC As Short

                            intRC = rows10.Length
                            varAC = 15
                            If intRC = 0 Then
                                varAC = 15
                            Else
                                For Count4 = 0 To intRC - 1
                                    num1 = NZ(rows10(Count4).Item("CONCENTRATION"), -1)
                                    'varNom is double, must convert to single
                                    If num1 = numNomConc Then
                                        var2 = NZ(rows10(Count4).Item("FLAGPERCENT"), 15)
                                        var4 = NZ(rows10(Count4).Item("ANALYTEFLAGPERCENT"), var2)
                                        varAC = CDec(NZ(var4, 15))
                                        varAC = NZ(varAC, 15)
                                        Exit For
                                    End If
                                Next
                            End If

                            '*****

                            'intRC = rows10.Length
                            'If intRC = 0 Then
                            '    varAC = 15
                            '    var5 = 1
                            'Else
                            '    'var4 = rows10(0).Item("FLAGPERCENT")
                            '    var4 = rows10(0).Item("ANALYTEFLAGPERCENT")
                            '    varAC = CDec(NZ(var4, 15))
                            '    varAC = NZ(varAC, 15)
                            'End If

                            'arrBCQCs(4, Count2) = varAC

                            'Dim arrFP(4, 20) 'flag percent array
                            arrFP(1, Count2) = varAC
                            arrFP(2, Count2) = varAC

                            ''****

                            Call SetHighAndLowCriteria(CDbl(var1), CDbl(varAC), CDbl(varAC), hi, lo)


                            'arrfp
                            '1=max, 2=min, 3=hi, 4=lo

                            'arrBCQCs(9, int1)
                            '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hi, 6=Lo, 7=#ofReplicates, 
                            '8=ASSAYID, 9=Aliquotfactor

                            arrBCQCs(5, Count2) = hi
                            arrBCQCs(6, Count2) = lo

                            arrBCQCs(9, Count2) = var6 'ALIQUOTFACTOR

                            arrFP(3, Count2) = hi
                            arrFP(4, Count2) = lo

                            '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel

                        End If

                    Next Count3


                    ctQCs = Count2

                    'find #ofReplicates for each level

                    'arrBCQCs(9, int1)
                    '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 
                    '8=ASSAYID, 9=Aliquotfactor

                    maxRep = 0
                    For Count2 = 1 To ctQCs
                        ' var2 = arrctQCs(2, Count1)
                        var2 = arrBCQCs(2, Count2) 'nomconc
                        'var1 = arrBCQCs(1, Count2) 'Levelnumber

                        'can't get LevelNumber from arrBCQC's anymore
                        'must get it directly from the AssayID


                        'strF = strFAssayID & " AND ASSAYLEVEL = " & var1 & " AND ANALYTEID = " & intAnalyteID
                        'filter by NomcConc instead

                        var6 = NZ(arrBCQCs(9, Count2), 1) 'aliquotfactor

                        strF = strFAssayID & " AND NOMCONC = " & var2 & " AND ANALYTEID = " & intAnalyteID & " AND ALIQUOTFACTOR = " & var6
                        Dim dvReps As DataView
                        Try
                            dvReps = New DataView(tblBCQCConcs, strF, "REPLICATENUMBER ASC", DataViewRowState.CurrentRows)
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try
                        Dim tblReps As DataTable = dvReps.ToTable("a", True, "REPLICATENUMBER")
                        int2 = tblReps.Rows.Count
                        If int2 = 0 Then
                        Else
                            var2 = tblReps.Rows(int2 - 1).Item("REPLICATENUMBER")
                            arrBCQCs(7, Count2) = var2
                            If var2 > maxRep Then
                                maxRep = var2
                            End If
                        End If
                    Next Count2

                    'find Interpolated QC Standard Concentrations for all analytical runs
                    'use existing arrBCQCConcs array because there's no difference in elements
                    ReDim arrBCQCConcs(9, 100)
                    '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=Hello, 8=Lo, 9=AssayID
                    Count2 = 0

                    'If boolIncludePSAE Then
                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNANALYTEREGRESSIONSTATUS <> 4"
                    'Else
                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNANALYTEREGRESSIONSTATUS <> 4"
                    'End If

                    strF = strFAssayID & " AND RUNTYPEID <> 3 AND ANALYTEID = " & intAnalyteID & " AND RUNANALYTEREGRESSIONSTATUS <> 4"

                    Erase drows

                    drows = tblBCQCConcs.Select(strF)
                    int1 = drows.Length

                    Dim strSName As String

                    Dim intAAA As Short
                    Dim strRRID As String
                    For Count2 = 0 To int1 - 1
                        For Count3 = 1 To ctQCs
                            If Count3 > int1 Then
                                Exit For
                            End If
                            If Count2 >= UBound(arrBCQCConcs, 2) Then
                                ReDim Preserve arrBCQCConcs(9, UBound(arrBCQCConcs, 2) + 100)
                            End If

                            'arrBCQCs(9, int1)
                            '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 
                            '8=ASSAYID, 9=Aliquotfactor

                            strRRID = arrBCQCs(3, Count3) 'RunID
                            arrBCQCConcs(1, Count2 + 1) = drows(Count2).Item("ASSAYLEVEL")
                            str2 = drows(Count2).Item("SAMPLENAME")
                            arrBCQCConcs(1, Count2 + 1) = Count3
                            If InStr(1, str2, strRRID, vbTextCompare) = 0 Then
                                num1 = 0
                                arrBCQCConcs(2, Count2 + 1) = num1 'concentration
                                arrBCQCConcs(3, Count2 + 1) = arrBCQCConcs(3, Count2) 'run id
                                arrBCQCConcs(4, Count2 + 1) = "Y" 'eliminated flag
                                arrBCQCConcs(5, Count2 + 1) = "NA" 'sample name
                                arrBCQCConcs(6, Count2 + 1) = 1 'aliquot factor
                                arrBCQCConcs(7, Count2 + 1) = 0 'Hi
                                arrBCQCConcs(8, Count2 + 1) = 0 'Lo
                                arrBCQCConcs(9, Count2 + 1) = 0
                            Else
                                num1 = NZ(drows(Count2).Item("CONCENTRATION"), 0)
                                'DO NOT SIGFIG HERE!
                                'DO IT LATER!
                                'If boolLUseSigFigs Then
                                '    num1 = SigFigOrDec(num2, LSigFig, False)
                                'Else
                                '    num1 = RoundToDecimalRAFZ(num2, LSigFig)
                                'End If

                                arrBCQCConcs(2, Count2 + 1) = num1
                                arrBCQCConcs(3, Count2 + 1) = drows(Count2).Item("RUNID")
                                arrBCQCConcs(4, Count2 + 1) = NZ(drows(Count2).Item("ELIMINATEDFLAG"), "N")
                                'arrBCQCConcs(5, Count2 + 1) = SampleName(CStr(NZ(drows(Count2).Item("SAMPLENAME"), "")))
                                strSName = NZ(drows(Count2).Item("SAMPLENAME"), "")
                                If boolSampleName01 Then
                                    strSName = SampleName(strSName)
                                End If
                                arrBCQCConcs(5, Count2 + 1) = strSName ' NZ(drows(Count2).Item("SAMPLENAME"), "")
                                arrBCQCConcs(6, Count2 + 1) = CDec(NZ(drows(Count2).Item("ALIQUOTFACTOR"), 1))
                                arrBCQCConcs(7, Count2 + 1) = arrBCQCs(5, Count3) 'Hi
                                arrBCQCConcs(8, Count2 + 1) = arrBCQCs(6, Count3) 'Lo
                                arrBCQCConcs(9, Count2 + 1) = drows(Count2).Item("ASSAYID")

                            End If

                        Next Count3

                    Next Count2

                    'inttemprows = arrAnalytes(7, Count1) '#accepted runs
                    int1 = (int1 * 3) - 1

                    '******
                    'arrBCQCConcs(9, 100)
                    '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=Hello, 8=Lo, 9=AssayID


                    Dim boolExRow As Boolean

                    'for each accepted analytical run
                    Count5 = 0
                    Dim int20 As Short
                    Dim intStep As Short
                    'need to find intstep

                    inttemprows = ctAnalyticalRuns 'ctAnalytical runs came from earlier

                    ReDim arrBCStdActual(10, ctAnalyticalRuns * maxRep * ctQCs)
                    '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 
                    '6=AliquotFactor(DilFactor), 7=Hello, 8=Lo, 9=AssayID, 10=FlagPercent
                    Erase DilQCFactor
                    ReDim DilQCFactor(ctQCs)

                    Dim nomConc As Decimal
                    Dim numRep As Short
                    Dim varAF
                    Dim varAF1
                    Dim varFP
                    Dim varQCAF
                    Dim varID

                    'For Count2 = 0 To int10 - 1 Step intRP 'step because tblRegCon has doublerow or triplerow entries
                    For Count2 = 0 To ctAnalyticalRuns - 1
                        'need maxRep rows for each accepted run
                        int20 = rowsAllRuns(Count2).Item("RUNID")

                        'For Count3 = 0 To maxRep - 1
                        'establish array going across table ctQC number of times
                        boolExRow = True
                        For Count4 = 1 To ctQCs
                            If Count4 = 4 Then
                                str1 = "aaa"
                            End If
                            '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel
                            var2 = arrBCQCs(1, Count4) '.Item("LevelNumber")
                            var3 = arrBCQCs(2, Count4) 'NomConc  CONCENTRATION
                            varAF1 = NZ(arrBCQCs(9, Count4), 1) 'aliquotfactor

                            nomConc = NZ(var3, 0)
                            'Count5 = Count5 + 1
                            'If boolIncludePSAE Then
                            '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            'Else
                            '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            'End If

                            'ignore PSAE
                            'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            'don't use ASSAYLEVEL anymore
                            '20160503 LEE: Aaack! MUST use ASSAYLEVEL. If QC-High and Diln have same NomConc, then get 2x too many reps
                            'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            If INTQCLEVELGROUP = 0 Then 'use assaylevel
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                                'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                '20171124 LEE:
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4 AND ALIQUOTFACTOR = " & varAF1
                            ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                                var3 = arrBCQCs(3, Count4) 'ID
                                var4 = arrBCQCs(10, Count4) 'QCLABEL  check
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND QCLABEL = '" & var3 & "' AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            Else
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            End If

                            Erase drowsF
                            Try
                                drowsF = tblBCQCConcs.Select(strF, "RUNSAMPLEORDERNUMBER ASC")
                            Catch ex As Exception
                                var1 = ex.Message
                                var1 = var1
                            End Try

                            intF = drowsF.Length

                            For Count5 = 0 To intF - 1

                                If intF = 0 Then 'enter base values'?????
                                    ''establish base values for array
                                    'arrBCStdActual(1, Count5) = 0 'LEVELNUMBER to designate NA
                                    'arrBCStdActual(2, Count5) = 0 'Conc
                                    ''arrBCStdActual(3, Count5) = drowsF(Count3).Item("RunID")
                                    'arrBCStdActual(4, Count5) = "Y" 'Eliminated Flag
                                    'arrBCStdActual(5, Count5) = "NA" 'SampleName

                                Else 'enter in total set of values

                                    Dim rowsz As DataRow = tblZ.NewRow
                                    rowsz.BeginEdit()

                                    boolExRow = False
                                    var5 = drowsF(Count5).Item("SampleName")
                                    var6 = drowsF(Count5).Item("AssayID")
                                    var7 = drowsF(Count5).Item("Concentration")
                                    varID = drowsF(Count5).Item("QCLABEL")
                                    'DON'T DO SIGFIGS HERE!
                                    'DO IT LATER IN THE CODE!!
                                    'If boolLUseSigFigs Then
                                    '    var7 = SigFigOrDec(NZ(var7, 0), LSigFig, False)
                                    'Else
                                    '    var7 = RoundToDecimalRAFZ(NZ(var7, 0), LSigFig)
                                    'End If

                                    varAF = NZ(drowsF(Count5).Item("AliquotFactor"), 1)

                                    Dim varg
                                    varg = drowsF(Count5).Item("RunID")

                                    'evaluate DilQCFactor

                                    'find actual hi/lo
                                    Dim rows10() As DataRow
                                    'strF = "CONCENTRATION = " & nomConc & " AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND ASSAYID = " & var6
                                    'strF = strFAssayID & "AND CONCENTRATION = " & varNom & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYID = " & var6
                                    '20171128 LEE:
                                    strF = strFAssayID & "AND CONCENTRATION = " & nomConc & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYID = " & var6

                                    'if Conc < 1, then the query return 0 records
                                    'must do something different
                                    'varnom =
                                    'varFP = GetANALYTEFLAGPERCENTAnova(varNom, varg, vAnalyteID, tblLevelCrit)
                                    varFP = GetANALYTEFLAGPERCENTAnova(nomConc, varg, vAnalyteID, tblLevelCrit)
                                    'done

                                    'varg =

                                    'calculate hi/lo
                                    Call SetHighAndLowCriteria(CDbl(nomConc), CDbl(varFP), CDbl(varFP), hi, lo)
                                    'end calculate hi/lo

                                    ''var1 might have to go back in
                                    'var1 = arrBCStdActual(6, Count5) 'aliquotfactor
                                    If varAF <> 1 Then
                                        If Len(DilQCFactor(Count4)) = 0 Then
                                            DilQCFactor(Count4) = varAF ' NZ(drowsF(Count5).Item("AliquotFactor"), 1) ' arrBCStdActual(6, Count5)
                                        End If
                                    End If
                                    'If InStr(1, var5, "Dil", vbTextCompare) > 0 Then
                                    '    If Len(DilQCFactor(Count4)) = 0 Then
                                    '        DilQCFactor(Count4) = arrBCStdActual(6, Count5)
                                    '    End If
                                    'End If

                                    'record nomConc

                                    rowsz.Item("NomConc") = nomConc
                                    rowsz.Item("Conc") = var7
                                    rowsz.Item("ALIQUOTFACTOR") = varAF ' NZ(drowsF(Count5).Item("AliquotFactor"), 1)
                                    var6 = NZ(drowsF(Count5).Item("EliminatedFlag"), "N")
                                    rowsz.Item("ELIMINATEDFLAG") = var6
                                    rowsz.Item("BOOLOUTLIER") = False
                                    rowsz.Item("HI") = hi 'arrBCQCs(5, Count4)
                                    rowsz.Item("LO") = lo 'arrBCQCs(6, Count4)
                                    rowsz.Item("RunID") = varg
                                    rowsz.Item("FlagPercent") = varFP 'arrBCQCs(4, Count4)
                                    rowsz.Item("numRep") = Count5 + 1
                                    If StrComp(var6, "Y", CompareMethod.Text) = 0 Then
                                        var7 = drowsF(Count5).Item("DECISIONREASON")
                                        rowsz.Item("DECISIONREASON") = NZ(var7, "")
                                    End If

                                    rowsz.Item("v1") = varFP
                                    rowsz.Item("v2") = varFP

                                    rowsz.Item("QCLABEL") = varID

                                    rowsz.Item("ASSAYLEVEL") = var2

                                    rowsz.EndEdit()
                                    tblZ.Rows.Add(rowsz)

                                End If
                            Next

                            '''''''''''''''''console.writeline(CStr(arrBCStdActual(2, Count5)))
                        Next Count4
                        'Next Count3

                    Next Count2

                    ''debug
                    'var1 = ChrW(9)
                    'For Count2 = 0 To tblZ.Columns.Count - 1
                    '    var2 = tblZ.Columns(Count2).ColumnName
                    '    var1 = var1 & ChrW(9) & var2
                    'Next
                    ''console.writeline(var1)

                    'For Count3 = 0 To tblZ.Rows.Count - 1
                    '    var1 = ChrW(9)
                    '    For Count2 = 0 To tblZ.Columns.Count - 1
                    '        var2 = tblZ.Rows(Count3).Item(Count2)
                    '        var1 = var1 & ChrW(9) & var2
                    '    Next
                    '    'console.writeline(var1)
                    'Next

                    '******
                    intLeg = 0
                    ctQCLegend = 0
                    ctDilLeg = 0

                    For Count2 = 1 To ctQCs
                        'var1 = NZ(DilQCFactor(Count2), "")
                        var2 = arrBCQCs(9, Count2) 'debug
                        var1 = NZ(arrBCQCs(9, Count2), 1) 'aliquotfactor
                        If var1 = 1 Then
                        Else
                            intLeg = intLeg + 1

                            ctDilLeg = ctDilLeg + 1
                            'configure first legend item
                            var4 = arrBCQCs(2, Count2)
                            'var4 = Format(arrBCQCs(2, Count2), "0")
                            'var2 = Sheets("AnalRefTables").Range("LLOQUnits").Offset(0, Count1).Value
                            'var3 = Format(1 / DilQCFactor(Count2), "0")
                            'var3 = Format(1 / var1, "0")
                            var3 = GetDilnFactor(CDec(var1)) '20190220 LEE
                            Dim strAN As String = GetAN(var3)

                            arrLegend(1, intLeg) = Chr(96 + intLeg) 'a,b,c,etc
                            'arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & var4 & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                            arrLegend(2, intLeg) = "Dilution QCs undiluted concentration " & var4 & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
                            arrLegend(3, intLeg) = True
                            arrLegend(4, intLeg) = True
                            arrQCLegend(1, intLeg) = Chr(96 + intLeg) 'a,b,c,etc
                            'arrQCLegend(2, intLeg) = "Dilution QCs undiluted concentration " & var4 & " " & strConcUnits & "; a 1:" & var3 & " dilution with blank matrix was performed prior to extraction and analysis."
                            arrQCLegend(2, intLeg) = "Dilution QCs undiluted concentration " & var4 & " " & strConcUnits & "; " & strAN & " " & var3 & "-fold dilution with blank matrix was performed prior to extraction and analysis."
                            arrQCLegend(3, intLeg) = True
                            arrQCLegend(4, intLeg) = True
                            ctQCLegend = ctQCLegend + 1
                        End If
                    Next
                    'intLeg = intLeg + 1
                    'arrLegend(1, intLeg) = Chr(96 + intLeg) '
                    'arrLegend(2, intLeg) = "Value outside of acceptance criteria but included in summary statistics."
                    'arrLegend(3, intLeg) = True
                    'arrQCLegend(1, intLeg) = Chr(96 + intLeg) '"a"
                    'arrQCLegend(2, intLeg) = "Value outside of acceptance criteria but included in summary statistics."
                    'arrQCLegend(3, intLeg) = True
                    'ctQCLegend = ctQCLegend + 1

                    'int1 = (ctAnalyticalRuns * (maxRep + 1)) + 9
                    'int1 = (ctAnalyticalRuns * (maxRep + 1)) + 4

                    int1 = 0
                    int1 = int1 + 3 'for header
                    int1 = int1 + 1 'for blank row
                    int1 = int1 + (ctAnalyticalRuns * (maxRep + 1))

                    'Increment for Statistics Sections
                    Dim intCSN As Short
                    intCSN = countNumStatsRows()
                    int1 = int1 + intCSN

                    If boolQCREPORTACCVALUES Then
                    Else


                        'Else
                        'ctExp = 8
                        'int1 = (ctAnalyticalRuns * (maxRep + 1)) + 17
                        'int1 = (ctAnalyticalRuns * (maxRep + 1)) + 12
                        int1 = int1 + 1 'subheader 1
                        int1 = int1 + 1 'blank row
                        int1 = int1 + 1 'subheader2
                        ctExp = 3

                        'Increment for Statistics Sections
                        int1 = int1 + intCSN
                        ctExp = ctExp + intCSN

                    End If

                    wrdSelection = wd.Selection()

                    Dim intCols As Short
                    If boolSTATSDIFFCOL Then
                        intCols = (ctQCs * 2) + 1
                    Else
                        intCols = ctQCs + 1
                    End If


                    Try

                        '20180913 LEE:
                        Call IncrNextTableNumber(wd)

                        If boolPlaceHolder Then
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=1, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        Else
                            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=int1, NumColumns:=intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                        End If

                        .Selection.Tables.Item(1).Select()

                        Call SetCellPaddingZero(.Selection.Tables.Item(1))

                        Call GlobalTableParaFormat(wd)

                        '20171220 LEE: Do not set table size, use the style default table
                        '.Selection.Font.Size = fontsize - 1
                        .Selection.Tables.Item(1).Cell(1, 1).Select()

                        .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                        .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                        '.Selection.Tables.Item(1).Columns.Item(1).Width = 86
                        'For Count2 = 1 To ctQCs
                        '    .Selection.Tables.Item(1).Columns.Item(Count2 + 1).Width = 50
                        'Next

                        .Selection.Tables.Item(1).Select()
                        .Selection.Rows.AllowBreakAcrossPages = False

                        removeBorderButLeaveTopAndBottom(wd)
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        '.Selection.Font.Size = 11

                        If boolPlaceHolder Then

                            .Selection.Tables.Item(1).Select()
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                            .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone


                            'strA = arrAnalytes(14, Count1)
                            If gNumMatrix = 1 Then
                                strA = strAnalC
                            Else
                                strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                            End If
                            'No. Now just send strAnal
                            strA = strAnal
                            strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                            Call EnterTableNumber(wd, strTName, 3, strA, strTempInfo, intTableID, intGroup, idTR)
                            'Note: strTName is byRef and will return Table, number, caption, label

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

                        .Selection.Tables.Item(1).Cell(1, 2).Select()
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        Try
                            .Selection.Cells.Merge()
                        Catch ex As Exception

                        End Try
                        '.Selection.Font.Size = 11
                        .Selection.Font.Bold = False
                        .Selection.TypeText(Text:="Nominal Concentrations")
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        'border top and bottom of range
                        .Selection.Tables.Item(1).Cell(1, 1).Select()
                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                        '.Selection.MoveRight Unit:=Microsoft.Office.Interop.Word.wdunits.word.wdunits.wdCharacter, Count:=ctQCs, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                        .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                        .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)
                        .Selection.Tables.Item(1).Cell(2, 2).Select()

                        Dim int11 As Short
                        If boolSTATSDIFFCOL Then
                            int11 = 2
                        Else
                            int11 = 1
                        End If

                        'arrBCQCs(9, int1)
                        '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 
                        '8=ASSAYID, 9=Aliquotfactor

                        For Count2 = 1 To ctQCs
                            var1 = arrBCQCs(3, Count2)
                            .Selection.TypeText(Text:=CStr(arrBCQCs(3, Count2)))
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=int11)
                        Next

                        .Selection.Tables.Item(1).Cell(3, 1).Select()
                        If BOOLINCLUDEDATE Then

                            'var1 = strRRID & ChrW(10) & strConcUnits
                            .Selection.Tables.Item(1).Cell(2, 1).Select()
                            .Selection.TypeText(strWRunId)
                            .Selection.Tables.Item(1).Cell(3, 1).Select()
                            '.Selection.TypeText("(Analysis Date)")
                            '20180420 LEE:
                            .Selection.TypeText("(" & GetAnalysisDateLabel(intTableID) & ")")
                        Else

                            int1 = InStr(strWRunId, " ", CompareMethod.Text)
                            If int1 = 0 Then
                                str2 = strWRunId
                            Else
                                str1 = Mid(strWRunId, 1, int1 - 1)
                                str2 = Mid(strWRunId, int1 + 1, Len(strWRunId))
                            End If

                            'var1 = strRRID & ChrW(10) & strConcUnits
                            If int1 = 0 Then
                            Else
                                .Selection.Tables.Item(1).Cell(2, 1).Select()
                                .Selection.TypeText(str1)
                            End If

                            .Selection.Tables.Item(1).Cell(3, 1).Select()
                            .Selection.TypeText(str2)

                        End If
                        .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                        For Count2 = 1 To ctQCs
                            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                            If ctQCs > 4 Then
                                .Selection.Font.Size = .Selection.Font.Size - 1
                            End If

                            num1 = arrBCQCs(2, Count2)
                            If boolLUseSigFigs Then
                                str1 = DisplayNum(SigFigOrDec(CDbl(num1), LSigFig, False), LSigFig, False) 'conc
                            Else
                                str1 = Format(CDbl(num1), GetRegrDecStr(LSigFig)) 'conc
                            End If

                            If LboolNomConcParen Then
                                If StrComp(strOrientation, "P", CompareMethod.Text) = 0 Then
                                    If boolSTATSDIFFCOL Then
                                        var1 = str1 & ChrW(10) & "(" & strConcUnits & ")"
                                    Else
                                        var1 = "(" & str1 & ChrW(160) & strConcUnits & ")"
                                    End If
                                Else
                                    var1 = "(" & str1 & ChrW(160) & strConcUnits & ")"
                                End If
                            Else
                                var1 = str1 & ChrW(160) & strConcUnits
                            End If

                            .Selection.TypeText(Text:=var1)

                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                            If boolSTATSDIFFCOL Then
                                .Selection.TypeText(Text:=ReturnDiffLabel)
                                If Count2 = ctQCs Then
                                Else
                                    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                End If
                            End If

                        Next
                        'enter dilution qc superscripts
                        Count3 = 0

                        '20160318 LEE:
                        'new order of superscripts performed earlier in code

                        '20171118 LEE:
                        'Goofy scenario: LI00016 - ACHN172001
                        'has 

                        Count3 = 0 'ctDilLeg + 1
                        For Count2 = ctDilLeg To 1 Step -1

                            Count3 = Count3 + 1 ' - 1

                            If int11 = 1 Then
                                .Selection.Tables.Item(1).Cell(2, ctQCs - Count2 + 2).Select()
                            Else
                                .Selection.Tables.Item(1).Cell(2, ((ctQCs - Count2) * int11) + 2).Select()
                            End If
                            .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCharacter)

                            ''''''wdd.visible = True

                            .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter)
                            '.Selection.TypeText(Text:=CStr(Chr(Count3 + 96))) '
                            .Selection.TypeText(" " & CStr(Chr(Count3 + 96))) '
                            'superscript the footnote
                            .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1, Word.WdMovementType.wdExtend)
                            .Selection.Font.Superscript = True
                            .Selection.Font.Size = 12
                            .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                        Next

                        'For Count2 = ctDilLeg - 1 To 0 Step -1
                        '    Count3 = Count3 + 1
                        '    '.Selection.Tables.Item(1).Cell(2, (ctQCs + 1) - (Count2 * int11)).Select()
                        '    .Selection.Tables.Item(1).Cell(2, ((ctQCs * int11)) - (Count2 * int11)).Select()
                        '    .Selection.Tables.Item(1).Cell(2, ((ctQCs + 1) * int11) - ((Count2 * int11) + int11)).Select()
                        '    .Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCharacter)
                        '    .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter)
                        '    '.Selection.TypeText(Text:=CStr(Chr(Count3 + 96))) '
                        '    .Selection.TypeText(" " & CStr(Chr(Count3 + 96))) '
                        '    'superscript the footnote
                        '    .Selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1, Word.WdMovementType.wdExtend)
                        '    .Selection.Font.Superscript = True
                        '    .Selection.Font.Size = 12
                        '    .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                        'Next

                        'autofit table
                        Call AutoFitTable(wd, False)

                        'enter concentration values

                        Dim intRow As Short

                        intRow = 4 'Counter for row selection' make one before
                        Count5 = 0 'Counter for arr
                        .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                        str1 = ""
                        str2 = ""
                        ctP = 1
                        ctLegend = 0
                        boolGo = True

                        intCS = Count4

                        ''''''''wdd.visible = True

                        'start new

                        Dim arrCS(1)
                        ReDim arrCS(ctQCs)
                        Dim boolDiffColOK As Boolean
                        Dim varAConc
                        Dim varAConcCorr
                        Dim intTCol As Short
                        Dim ctNA As Short

                        ctP = 0

                        For Count2 = 0 To ctAnalyticalRuns - 1

                            intRow = intRow + 1

                            frmH.lblProgress.Text = strM1 & ChrW(10) & "For Analytical Run # " & Count2 + 1 & " of " & ctAnalyticalRuns & "..."
                            frmH.Refresh()

                            'need maxRep rows for each accepted run
                            'int20 = CInt(drows(Count2).Item("RUNID"))
                            'int20 = CInt(tblTT.Rows(Count2).Item("RUNID"))
                            int20 = rowsAllRuns(Count2).Item("RUNID")

                            var1 = int20 ' arrBCStdActual(3, Count5)
                            var10 = var1

                            .Selection.Tables.Item(1).Cell(intRow, 1).Select()

                            .Selection.TypeText(Text:=CStr(int20))
                            '''wdd.visible = True
                            If BOOLINCLUDEDATE Then
                                Dim intRR As Int32
                                .Selection.Tables.Item(1).Cell(intRow + 1, 1).Select()
                                str1 = GetDateFromRunID(NZ(int20, 0), LDateFormat, intGroup, idTR)
                                .Selection.TypeText("(" & str1 & ")")
                                .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                            End If

                            '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, Count:=1)

                            Dim arrRep(5, maxRep)
                            '1=Row#,2=QC#, 3=Deleted (Yes/No)

                            For Count3 = 0 To maxRep - 1
                                'establish array going across table ctQC number of times
                                boolExRow = True
                                intTCol = 1
                                ctNA = 0
                                For Count4 = 1 To ctQCs

                                    intTCol = intTCol + 1

                                    '.Selection.Tables.Item(1).Cell(intRow, intTCol).Select()

                                    If Count4 = 4 Then
                                        str1 = "aaa"
                                    End If
                                    '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel
                                    var2 = arrBCQCs(1, Count4) '.Item("LevelNumber")
                                    var3 = arrBCQCs(2, Count4) 'CONCENTRATION
                                    var4 = NZ(arrBCQCs(9, Count4), 1) 'aliquot factor
                                    var5 = arrBCQCs(3, Count4) 'QC label
                                    nomConc = var3

                                    '20160318 LEE: This filter is incorrect
                                    'it needs to also take in to account aliquot factor
                                    'strF = "NomConc = " & var3 & " AND RunID = " & int20 & " AND numRep = " & Count3 + 1


                                    If INTQCLEVELGROUP = 0 Then 'use assaylevel
                                        strF = "ASSAYLEVEL = " & var2 & " AND ALIQUOTFACTOR = " & NZ(var4, 1) & " AND RunID = " & int20 & " AND numRep = " & Count3 + 1
                                        'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                                        strF = "NOMCONC = " & var3 & " AND ALIQUOTFACTOR = " & NZ(var4, 1) & " AND RunID = " & int20 & " AND numRep = " & Count3 + 1
                                    ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                                        strF = "QCLABEL = '" & var5 & "' AND ALIQUOTFACTOR = " & NZ(var4, 1) & " AND RunID = " & int20 & " AND numRep = " & Count3 + 1
                                    Else
                                        strF = "ASSAYLEVEL = " & var2 & " AND ALIQUOTFACTOR = " & NZ(var4, 1) & " AND RunID = " & int20 & " AND numRep = " & Count3 + 1
                                    End If

                                    Dim rowT() As DataRow
                                    rowT = tblZ.Select(strF, "numRep ASC")
                                    intF = rowT.Length

                                    'If boolIncludePSAE Then
                                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0"
                                    'Else
                                    '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3"
                                    'End If
                                    'Erase drowsF
                                    'drowsF = tblBCQCConcs.Select(str1, "RUNSAMPLEORDERNUMBER ASC")
                                    'intF = drowsF.Length

                                    arrRep(1, Count3 + 1) = intRow
                                    arrRep(3, Count3 + 1) = "No"

                                    If intF = 0 And boolQCNA = False And Count4 > 1 Then
                                    Else
                                        .Selection.Tables.Item(1).Cell(intRow, intTCol).Select()
                                    End If

                                    If intF = 0 Then 'enter base values
                                        If boolQCNA Then
                                            .Selection.TypeText("NA")
                                        End If

                                        If boolSTATSDIFFCOL Then
                                            intTCol = intTCol + 1
                                            If boolQCNA Then
                                                .Selection.Tables.Item(1).Cell(intRow, intTCol).Select()
                                                .Selection.TypeText("NA")
                                            End If

                                        End If

                                        ctNA = ctNA + 1

                                        '1=Row#,2=QC#, 3=Deleted (Yes/No)

                                        If Count4 = ctQCs Then

                                            boolOC = False

                                            If ctNA = ctQCs Then

                                                boolOC = True
                                                'remove this row
                                                For Count5 = 1 To intCols
                                                    .Selection.Tables.Item(1).Cell(intRow, Count5).Select()
                                                    .Selection.Delete()
                                                Next
                                                intRow = intRow - 1
                                                arrRep(3, Count3 + 1) = "Yes"
                                                Dim intFRow
                                                If Count3 = maxRep - 1 Then
                                                    'evaluate reps
                                                    Dim boolNo As Boolean = False
                                                    For Count5 = 1 To maxRep
                                                        If Count5 = 1 Then
                                                            intFRow = arrRep(1, Count5)
                                                        End If
                                                        var1 = NZ(arrRep(3, Count5), "")
                                                        If StrComp(var1, "Yes", CompareMethod.Text) = 0 Then
                                                        Else
                                                            boolNo = True
                                                            Exit For
                                                        End If
                                                    Next
                                                    If boolNo Then
                                                        're-enter RunId
                                                        .Selection.Tables.Item(1).Cell(intFRow, 1).Select()
                                                        .Selection.TypeText(Text:=CStr(int20))
                                                    End If

                                                End If
                                                'If Count3 = maxRep - 1 Then
                                                '    intRow = intRow - 2
                                                'Else
                                                '    'replace runid
                                                '    .Selection.Tables.Item(1).Cell(intRow, 1).Select()
                                                '    .Selection.TypeText(Text:=CStr(int20))
                                                '    intRow = intRow - 1
                                                'End If
                                            End If
                                        End If

                                    Else '

                                        varConc = rowT(0).Item("Conc")

                                        varAConc = NZ(rowT(0).Item("Conc"), 0)
                                        var2 = NZ(rowT(0).Item("ALIQUOTFACTOR"), 1)
                                        var3 = varAConc / var2
                                        If boolLUseSigFigs Then
                                            varAConcCorr = SigFigOrDec(var3, LSigFig, False)
                                        Else
                                            varAConcCorr = RoundToDecimalRAFZ(var3, LSigFig)
                                        End If

                                        var1 = rowT(0).Item("ELIMINATEDFLAG") ' arrBCStdActual(4, Count5) 'Error Flag
                                        boolDiffColOK = True

                                        If IsDBNull(varConc) Then

                                            boolOC = True

                                            boolDiffColOK = False
                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            'strA = ChrW(intLeg + intLegStart) 'Chr(96 + intLeg)
                                            strA = Chr(intLeg + intLegStart) 'Chr(96 + intLeg)

                                            var1 = "Y"

                                            '******

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            'Dim var6
                                            var6 = "No Value: " & NZ(rowT(0).Item("DECISIONREASON"), "No reason recorded.")
                                            var5 = CDec(NZ(rowT(0).Item("FlagPercent"), 15))

                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(var5, var5, vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            '******

                                            'fontsize = .Selection.Font.Size

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If


                                            .Selection.TypeText(Text:="NV")

                                            'arrCS(intCS) = var4
                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)

                                        ElseIf StrComp(var1, "Y", vbTextCompare) = 0 Then

                                            boolOC = True

                                            boolDiffColOK = False
                                            intExp = intExp + 1
                                            intLeg = intLeg + 1
                                            'strA = ChrW(intLeg + intLegStart) 'Chr(96 + intLeg)
                                            strA = Chr(intLeg + intLegStart) 'Chr(96 + intLeg)
                                            'str1 = "Dilution QC failed to meet acceptance criteria. Excluded from summary statistics."
                                            'str1 = "Value outside of acceptance criteria (" & RoundToDecimal(arrFP(Count3), 0) & "% theoretical) and excluded from summary statistics because the value is a statistical outlier according to the [OUTLIERMETHOD]."

                                            'str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(rowT(0).Item("FlagPercent"), 0) & "% theoretical) and excluded from summary statistics because the value is a statistical outlier according to the " & ReturnOutlierMethod() & "."

                                            '******

                                            '20160305 LEE:
                                            'Added DECISIONREASON code
                                            'Dim var6
                                            var6 = NZ(rowT(0).Item("DECISIONREASON"), "No reason recorded.") ' rowT(0).Item("DECISIONREASON")
                                            var5 = CDec(NZ(rowT(0).Item("FlagPercent"), 15))
                                            'Set Legend String
                                            str1 = GetLegendStringExcluded(var5, var5, vU, var6, intTableID, True, "")
                                            'Add to Legend Array
                                            ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                            '******

                                            'fontsize = .Selection.Font.Size

                                            If boolRedBoldFont Then
                                                .Selection.Font.Bold = True
                                                .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                            End If


                                            If boolLUseSigFigs Then
                                                var4 = CStr(DisplayNum(SigFigOrDecString(NZ(varAConcCorr, 0), LSigFig, False), LSigFig, False))
                                            Else
                                                var4 = CStr(Format(NZ(varAConcCorr, 0), GetRegrDecStr(LSigFig)))
                                            End If

                                            .Selection.TypeText(Text:=var4)
                                            'arrCS(intCS) = var4
                                            Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            '.Selection.TypeText Text:="NR"

                                        Else

                                            'determine if value is outside acceptance criteria
                                            var1 = NZ(varAConcCorr, 0)
                                            var2 = rowT(0).Item("Hi") ' arrBCStdActual(7, Count5) 'Hello
                                            var3 = rowT(0).Item("Lo") 'arrBCStdActual(8, Count5) 'Lo
                                            hi = var2
                                            lo = var3
                                            'If var1 > hi Or var1 < lo Then 'flag

                                            'Note: this sub does not contain assigned samples
                                            v1 = rowT(0).Item("FlagPercent")
                                            v1 = NZ(v1, 15)
                                            v1 = CDec(v1)
                                            v2 = v1
                                            vU = 0
                                            varNom = rowT(0).Item("NomConc")


                                            If OutsideAccCrit(var1, varNom, v1, v2, NZ(vU, 0)) Then


                                                'arrfp
                                                '1=max, 2=min, 3=hi, 4=lo

                                                intLeg = intLeg + 1
                                                'strA = ChrW(intLeg + intLegStart)
                                                strA = Chr(intLeg + intLegStart)
                                                'str1 = "Value outside of acceptance criteria (" & RoundToDecimal(arrFP(Count3), 0) & "% theoretical) but included in summary statistics."
                                                'str1 = "Value outside of acceptance criteria (" & RoundToDecimal(arrFP(Count3), 0) & "% theoretical) and excluded from summary statistics because the value is a statistical outlier according to the [OUTLIERMETHOD]."
                                                str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & v1 & "% theoretical) but included in summary statistics."

                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                'fontsize = .Selection.Font.Size

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                If boolLUseSigFigs Then
                                                    var4 = CStr(DisplayNum(SigFigOrDecString(NZ(varAConcCorr, 0), LSigFig, False), LSigFig, False))
                                                Else
                                                    var4 = CStr(Format(NZ(RoundToDecimalRAFZ(varAConcCorr, LSigFig), 0), GetRegrDecStr(LSigFig)))
                                                End If

                                                .Selection.TypeText(Text:=var4)
                                                'arrCS(intCS) = var4
                                                Call typeInSuperscriptFontSize12WithSpace(wd, strA)
                                            Else
                                                If boolLUseSigFigs Then
                                                    var4 = CStr(DisplayNum(SigFigOrDecString(NZ(varAConcCorr, 0), LSigFig, False), LSigFig, False))
                                                Else
                                                    var4 = CStr(Format(NZ(RoundToDecimalRAFZ(varAConcCorr, LSigFig), 0), GetRegrDecStr(LSigFig)))
                                                End If

                                                .Selection.TypeText(Text:=var4)
                                                'arrCS(intCS) = var4
                                            End If

                                        End If

                                        If boolSTATSDIFFCOL Then

                                            intTCol = intTCol + 1

                                            .Selection.Tables.Item(1).Cell(intRow, intTCol).Select()

                                            If IsNumeric(var4) And boolDiffColOK Then

                                                var2 = CalcREPercent(varAConcCorr, nomConc, intQCDec)
                                                var3 = Format(var2, strQCDec)


                                                If boolTHEORETICAL Then
                                                    If IsNumeric(var3) Then
                                                        var4 = 100 - var3
                                                        var3 = var4
                                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count4, "Accuracy", CSng(var4), CSng(var10), Count1, strDo, v1, v2, boolOC)
                                                    End If
                                                Else

                                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, Count4, "Accuracy", CSng(var3), CSng(var10), Count1, strDo, v1, v2, boolOC)

                                                End If

                                                '20180430 LEE:
                                                'must account for Endogenous Cmpds, NomConc = 0
                                                If IsNumeric(var3) Then
                                                    var3 = NomConcZero(nomConc, var3)
                                                End If

                                            Else

                                                If boolQCNA Then
                                                    var3 = "NA"
                                                Else
                                                    var3 = ""
                                                End If

                                            End If
                                            .Selection.TypeText(Text:=var3)

                                        End If

                                    End If


                                Next Count4

                                If Count2 = ctAnalyticalRuns - 1 And Count3 = maxRep - 1 Then
                                    'intRow = intRow - 2
                                Else
                                    intRow = intRow + 1
                                End If
                                'intRow = intRow + 1

                            Next Count3

                        Next Count2

                        'end new


                        'begin doing statistics
                        int1 = intRow ' .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
                        int1 = int1 + 1

                        '''''''''''''wdd.visible = True

                        .Selection.Tables.Item(1).Cell(int1 + 1, 1).Select()
                        If boolQCREPORTACCVALUES Then
                        Else
                            If intExp = 0 Then
                            Else
                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                '.Selection.Tables.Item(1).Cell(int1 + 1, 1).Select()
                                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                'enter some blank spaces to fool PageBreak function
                                '.selection.typetext(Text:="  ")
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
                                .Selection.Tables.Item(1).Cell(int1 + 2, 1).Select()
                            End If
                        End If

                        '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                        .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                        int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)
                        '.Selection.Tables.item(1).Cell(int1 + 1, 1).Select()
                        'Mean, SD, %CV, %Bias
                        int8 = 0

                        Call typeStatsLabels(wd, int8, int1 - 1, 1, False)

                        '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)


                        ''''''''''wdd.visible = True

                        '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, int8 - 1)
                        '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                        .Selection.Tables.Item(1).Cell(int1 + int8 - 1, 2).Select()

                        '***From BCStds
                        strM = "Entering Interpolated QC Concentration Statistics For " & arrAnalytes(1, Count1)
                        strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                        frmH.lblProgress.Text = strM
                        frmH.Refresh()
                        Dim intN As Short
                        Dim int12 As Short
                        int12 = 0
                        For Count3 = 1 To ctQCs * int11 Step int11

                            strM = "Entering Interpolated QC Concentration Statistics For " & arrAnalytes(1, Count1) & " for Level " & int12 + 1 & " of " & ctQCs & " calibration stds..."
                            strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                            frmH.lblProgress.Text = strM
                            frmH.Refresh()
                            intN = 0
                            int12 = int12 + 1

                            '20170718 LEE: WIL Sample Analysis study 313096 has a the same dilution level for two different aliquot factors
                            'must add aliquot factor to the queries below

                            '20160318 LEE: This filter is incorrect
                            'it needs to also take in to account aliquot factor
                            'strF = "NomConc = " & var3 & " AND RunID = " & int20 & " AND numRep = " & Count3 + 1

                            '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel


                            var2 = arrBCQCs(1, int12) '.Item("LevelNumber")
                            var3 = arrBCQCs(2, int12) 'CONCENTRATION  nomConc
                            varID = arrBCQCs(3, int12) 'ID
                            var4 = arrBCQCs(10, int12) 'QCLABEL check
                            var5 = NZ(arrBCQCs(9, int12), 1) 'aliquot factor

                            'If boolIncludePSAE Then
                            '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            'Else
                            '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            'End If
                            'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3
                            'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & Count3

                            ''ignore PSAE
                            ''strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            ''must have assaylevel, sometimes different diln have same nomconc
                            'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"

                            'intqclevelgroup
                            If INTQCLEVELGROUP = 0 Then 'use assaylevel
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                                var3 = arrBCQCs(3, int12) 'ID
                                var4 = arrBCQCs(10, int12) 'QCLABEL  check
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND QCLABEL = '" & var3 & "' AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            Else
                                strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                            End If

                            strF = strF & " AND ALIQUOTFACTOR = " & var5

                            Erase drows
                            drows = tblBCQCConcs.Select(strF)
                            int2 = drows.Length
                            ReDim arrBCQCActual(int2)

                            For Count5 = 0 To int2 - 1

                                'num1 = NZ(drows(Count3).Item("CONCENTRATION"), 0)
                                'num1 = SigFigOrDec(CDec(num1), LSigFig, True)
                                'frmh.arrBCStdConcs(2, Count3 + 1) = num1
                                'frmh.arrBCStdConcs(3, Count3 + 1) = drows(Count3).Item("RUNID")
                                'frmh.arrBCStdConcs(4, Count3 + 1) = drows(Count3).Item("ELIMINATEDFLAG")
                                var1 = NZ(drows(Count5).Item("ELIMINATEDFLAG"), "N") 'frmh.arrBCStdConcs(4, Count5)
                                If StrComp(var1, "Y", vbTextCompare) = 0 Or IsDBNull(drows(Count5).Item("CONCENTRATION")) Then 'exclude value
                                Else
                                    intN = intN + 1
                                    num1 = NZ(drows(Count5).Item("CONCENTRATION"), 0)
                                    num2 = NZ(drows(Count5).Item("ALIQUOTFACTOR"), 1)
                                    num3 = CDbl(num1 / num2)
                                    'num3 = SigFigOrDec(num3, LSigFig, False)
                                    If boolLUseSigFigs Then
                                        num4 = SigFigOrDecString(num3, LSigFig, False)
                                    Else
                                        num4 = RoundToDecimalRAFZ(num3, LSigFig)
                                    End If

                                    arrBCQCActual(intN) = num4
                                    'var7 = frmh.arrBCStdConcs(2, Count5)
                                End If
                            Next
                            'determine Sum
                            numSum = 0
                            If boolLUseSigFigs Then
                                numMean = SigFigOrDec(Mean(intN, arrBCQCActual), LSigFig, False)
                                numSD = SigFigOrDec(StdDev(intN, arrBCQCActual), LSigFig, False)
                            Else
                                numMean = RoundToDecimalRAFZ(Mean(intN, arrBCQCActual), LSigFig)
                                numSD = RoundToDecimalRAFZ(StdDev(intN, arrBCQCActual), LSigFig)
                            End If

                            '***End BCStds



                            int8 = 0
                            .Selection.Tables.Item(1).Cell(int1, Count3 + 1).Select()

                            ''''''''wdd.visible = True

                            varNom = CDec(arrBCQCs(2, int12))

                            hi = arrFP(3, int12)
                            lo = arrFP(4, int12)
                            v1 = arrFP(1, int12)
                            v2 = arrFP(2, int12)


                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Mean", numMean, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSMEAN Then
                                Try


                                    'record mean
                                    int8 = int8 + 1

                                    'determine if value is outside acceptance criteria
                                    'arrfp
                                    '1=max, 2=min, 3=hi, 4=lo

                                    'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                    If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                        intLeg = intLeg + 1
                                        strA = Chr(intLeg + intLegStart)
                                        str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(v1, 0) & "% theoretical) but included in summary statistics."
                                        'str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(arrFP(1, int12), 0) & "% theoretical)."

                                        'Add to Legend Array
                                        ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                        If boolRedBoldFont Then
                                            .Selection.Font.Bold = True
                                            .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                        End If

                                        '.Selection.TypeText(Text:=CStr(numMean))
                                        'Note: numMean has had sigfigs applied to it already
                                        If boolLUseSigFigs Then
                                            .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                        Else
                                            .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                        End If

                                        fonts = .Selection.Font.Size
                                        .Selection.Font.Superscript = True
                                        'fontsize = .Selection.Font.Size
                                        .Selection.Font.Size = 12
                                        '.Selection.TypeText(strA)
                                        .Selection.TypeText(" " & strA)
                                        .Selection.Font.Superscript = False
                                        .Selection.Font.Size = fonts
                                        'boolEnterDiff = True
                                    Else
                                        '.Selection.TypeText(Text:=CStr(numMean))
                                        If boolLUseSigFigs Then
                                            .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                        Else
                                            .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                        End If

                                        'boolEnterDiff = True
                                    End If

                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()
                                Catch ex As Exception

                                End Try
                            End If


                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "SD", numSD, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSSD Then
                                Try
                                    'record SD
                                    int8 = int8 + 1
                                    If intN < gSDMax Then
                                        .Selection.TypeText("NA")
                                    Else
                                        If boolLUseSigFigs Then
                                            .Selection.TypeText(Text:=CStr(DisplayNum(SigFigOrDec(numSD, LSigFig, False), LSigFig, False)))
                                        Else
                                            .Selection.TypeText(Text:=CStr(Format(numSD, GetRegrDecStr(LSigFig))))
                                        End If


                                    End If
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()
                                Catch ex As Exception

                                End Try
                            End If



                            Try
                                If intN < gSDMax Then
                                Else
                                    numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                    Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Precision", numPrec, CSng(var10), Count1, strDo, 0, 0, False)
                                End If

                            Catch ex As Exception

                            End Try
                            If boolSTATSCV Then
                                Try
                                    'record %CV
                                    int8 = int8 + 1
                                    If intN < gSDMax Then
                                        .Selection.TypeText("NA")
                                    Else

                                        .Selection.TypeText(Format(numPrec, strQCDec))

                                    End If
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()
                                Catch ex As Exception

                                End Try
                            End If


                            If boolSTATSBIAS And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)

                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, varNom, int12, 0, False)
                            End If

                            If boolSTATSBIAS And boolSTATSMEAN Then
                                Try
                                    'record %Bias
                                    int8 = int8 + 1

                                    If intN = 0 Then
                                        .Selection.TypeText("NA")
                                    Else
                                        '.Selection.TypeText(Format(numBias, strQCDec))
                                        '20180430 LEE:
                                        'must account for Endogenous Cmpds, NomConc = 0
                                        .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                    End If

                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()

                                Catch ex As Exception

                                End Try
                            End If


                            If boolTHEORETICAL And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                    numTheor = 100 + CDec(numTheor)

                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", numTheor, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numTheor = GetBiasFromDiffCol(idTR, varNom, int12, 0, False)
                                numTheor = 100 + CDec(numBias)
                            End If
                            If boolTHEORETICAL And boolSTATSMEAN Then
                                Try
                                    'record %Bias
                                    int8 = int8 + 1

                                    If intN = 0 Then
                                        .Selection.TypeText("NA")
                                    Else
                                        '.Selection.TypeText(Format(numTheor, strQCDec))
                                        '20180430 LEE:
                                        'must account for Endogenous Cmpds, NomConc = 0
                                        .Selection.TypeText(Text:=NomConcZero(varNom, numTheor))
                                    End If

                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()

                                Catch ex As Exception

                                End Try
                            End If

                            If boolSTATSDIFF And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)

                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, varNom, int12, 0, False)
                            End If
                            If boolSTATSDIFF And boolSTATSMEAN Then
                                Try
                                    'record %Bias
                                    int8 = int8 + 1

                                    '.Selection.TypeText(Format(numBias, strQCDec))
                                    '20180430 LEE:
                                    'must account for Endogenous Cmpds, NomConc = 0
                                    .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()

                                Catch ex As Exception

                                End Try
                            End If


                            If BOOLSTATSRE And (boolSTATSDIFFCOL = False Or (boolSTATSDIFFCOL And BOOLDIFFCOLSTATS = False)) Then
                                Try
                                    numBias = CalcREPercent(numMean, varNom, intQCDec)

                                    If intN = 0 Then
                                    Else
                                        Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", numBias, CSng(var10), Count1, strDo, 0, 0, False)
                                    End If

                                Catch ex As Exception

                                End Try
                            Else
                                'get numbias from average of %Bias columns
                                numBias = GetBiasFromDiffCol(idTR, varNom, int12, 0, False)
                            End If
                            If BOOLSTATSRE And boolSTATSMEAN Then
                                Try
                                    'record %RE
                                    int8 = int8 + 1

                                    '.Selection.TypeText(Format(numBias, strQCDec))
                                    '20180430 LEE:
                                    'must account for Endogenous Cmpds, NomConc = 0
                                    .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                    '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                    .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()

                                Catch ex As Exception

                                End Try
                            End If



                            Try
                                Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "n", intN, CSng(var10), Count1, strDo, 0, 0, False)
                            Catch ex As Exception

                            End Try
                            If boolSTATSN Then
                                Try
                                    'record n
                                    int8 = int8 + 1
                                    .Selection.TypeText(Text:=CStr(intN))
                                    '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                    '.Selection.Tables.Item(1).Cell(int1 + int8, 1).Select()


                                Catch ex As Exception

                                End Try
                            End If

                            If Count3 >= ctQCs * int11 Then
                            Else
                                '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)
                                '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, int8 - 1)
                                '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                            End If

                        Next

                        '******

                        If boolQCREPORTACCVALUES Then
                        Else
                            If intExp = 0 Then
                            Else
                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                .Selection.Tables.Item(1).Cell(int1 + int8 + 1, 1).Select()
                                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                                'enter some blank spaces to fool PageBreak function
                                '.selection.typetext(Text:="  ")
                                .Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                '.Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
                                .Selection.TypeText(Text:="Summary Statistics Including Outlier Values")
                                '.Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
                                .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                                Try
                                    .Selection.Cells.Merge()
                                Catch ex As Exception

                                End Try
                                '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)
                                With .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                                    .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                                End With

                                '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                .Selection.Tables.Item(1).Cell(int1 + int8 + 2, 1).Select()
                                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)

                                int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdStartOfRangeRowNumber)

                                '.Selection.Tables.item(1).Cell(int1 + 1, 1).Select()
                                'Mean, SD, %CV, %Bias
                                int8 = 0

                                Call typeStatsLabels(wd, int8, int1 - 1, 1, False)

                                '''''''wdd.visible = True

                                '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)
                                '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, int8 - 1)
                                '.Selection.Tables.Item(1).Cell(int1, 1).Select()
                                '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)

                                '***From BCStds
                                strM = "Entering Interpolated QC Concentration Statistics For " & strAnalC
                                strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                                frmH.lblProgress.Text = strM
                                frmH.Refresh()
                                int12 = 0
                                For Count3 = 1 To ctQCs * int11 Step int11

                                    int12 = int12 + 1
                                    strM = "Entering Interpolated QC Concentration Statistics For " & arrAnalytes(1, Count1) & " for Level " & int12 & " of " & ctQCs & " calibration stds..."
                                    strM = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                                    frmH.lblProgress.Text = strM
                                    frmH.Refresh()
                                    intN = 0

                                    '1=LevelNumber, 2=NomConcentration, 3=ID, 4=FlagPercent, 5=Hello, 6=Lo, 7=#ofReplicates, 8=ASSAYID, 9=AliquotFactor, 10=QCLabel
                                    var2 = arrBCQCs(1, int12) '.Item("LevelNumber")
                                    var3 = arrBCQCs(2, int12) 'CONCENTRATION 'nomconc
                                    var5 = NZ(arrBCQCs(9, int12), 1) 'aliquot factor
                                    If boolIncludePSAE Then
                                        str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    Else
                                        str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    End If
                                    'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3
                                    'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & Count3

                                    'ignore PSAE
                                    'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    'need assay level because sometimes diln have same nomconc
                                    strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"

                                    'intqclevelgroup
                                    If INTQCLEVELGROUP = 0 Then 'use assaylevel
                                        strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                        'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                                        strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND NOMCONC = " & var3 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                                        var3 = arrBCQCs(3, int12) 'ID
                                        var4 = arrBCQCs(10, int12) 'QCLABEL  check
                                        strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND QCLABEL = '" & var3 & "' AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    Else
                                        strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                                    End If

                                    strF = strF & " AND ALIQUOTFACTOR = " & var5

                                    Erase drows
                                    drows = tblBCQCConcs.Select(strF)
                                    int2 = drows.Length
                                    ReDim arrBCQCActual(int2)
                                    For Count5 = 0 To int2 - 1
                                        'num1 = NZ(drows(Count3).Item("CONCENTRATION"), 0)
                                        'num1 = SigFigOrDec(CDec(num1), LSigFig, True)
                                        'frmh.arrBCStdConcs(2, Count3 + 1) = num1
                                        'frmh.arrBCStdConcs(3, Count3 + 1) = drows(Count3).Item("RUNID")
                                        'frmh.arrBCStdConcs(4, Count3 + 1) = drows(Count3).Item("ELIMINATEDFLAG")
                                        var1 = NZ(drows(Count5).Item("ELIMINATEDFLAG"), "N") 'frmh.arrBCStdConcs(4, Count5)
                                        'If StrComp(var1, "Y", vbTextCompare) = 0 Or IsDBNull(drows(Count5).Item("CONCENTRATION")) Then 'exclude value
                                        'Else
                                        intN = intN + 1
                                        num1 = NZ(drows(Count5).Item("CONCENTRATION"), 0)
                                        num2 = NZ(drows(Count5).Item("ALIQUOTFACTOR"), 1)
                                        num3 = CDbl(num1 / num2)
                                        'num1 = SigFigOrDec(num3, LSigFig, False)
                                        If boolLUseSigFigs Then
                                            num4 = SigFigOrDecString(num3, LSigFig, False)
                                        Else
                                            num4 = RoundToDecimalRAFZ(num3, LSigFig)
                                        End If

                                        arrBCQCActual(intN) = num4
                                        'var7 = frmH.arrBCStdConcs(2, Count5)
                                        'End If
                                    Next
                                    'determine Sum
                                    numSum = 0
                                    If boolLUseSigFigs Then
                                        numMean = SigFigOrDec(Mean(intN, arrBCQCActual), LSigFig, False)
                                        numSD = SigFigOrDec(StdDev(intN, arrBCQCActual), LSigFig, False)
                                    Else
                                        numMean = RoundToDecimalRAFZ(Mean(intN, arrBCQCActual), LSigFig)
                                        numSD = RoundToDecimalRAFZ(StdDev(intN, arrBCQCActual), LSigFig)
                                    End If

                                    '***End BCStds

                                    int8 = 0
                                    .Selection.Tables.Item(1).Cell(int1, Count3 + 1).Select()

                                    varNom = CDec(arrBCQCs(2, int12))

                                    hi = arrFP(3, int12)
                                    lo = arrFP(4, int12)
                                    v1 = arrFP(1, int12)
                                    v2 = arrFP(2, int12)

                                    If boolSTATSMEAN Then
                                        'numMean = 0
                                        Try
                                            'record mean
                                            int8 = int8 + 1

                                            'determine if value is outside acceptance criteria
                                            'arrfp
                                            '1=max, 2=min, 3=hi, 4=lo

                                            'If (numMean > hi Or numMean < lo) And boolFootNoteQCMean Then 'flag
                                            If (OutsideAccCrit(numMean, varNom, v1, v2, NZ(vU, 0))) And boolFootNoteQCMean Then 'flag
                                                intLeg = intLeg + 1
                                                strA = Chr(intLeg + intLegStart)
                                                str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(v1, 0) & "% theoretical) but included in summary statistics."
                                                'str1 = "Value outside of acceptance criteria (" & ChrW(177) & " " & RoundToDecimal(arrFP(1, int12), 0) & "% theoretical)."

                                                'Add to Legend Array
                                                ctLegend = ctLegend + SetLegendArray(arrLegend, intLeg, str1, strA, True)

                                                If boolRedBoldFont Then
                                                    .Selection.Font.Bold = True
                                                    .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                                End If

                                                '.Selection.TypeText(Text:=CStr(numMean))
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If

                                                fonts = .Selection.Font.Size
                                                .Selection.Font.Superscript = True
                                                'fontsize = .Selection.Font.Size
                                                .Selection.Font.Size = 12
                                                '.Selection.TypeText(strA)
                                                .Selection.TypeText(" " & strA)
                                                .Selection.Font.Superscript = False
                                                .Selection.Font.Size = fonts
                                                'boolEnterDiff = True
                                            Else
                                                '.Selection.TypeText(Text:=CStr(numMean))
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(numMean, LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(numMean, GetRegrDecStr(LSigFig))))
                                                End If

                                                'boolEnterDiff = True
                                            End If
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If
                                    If boolSTATSSD Then
                                        Try
                                            'record SD
                                            int8 = int8 + 1
                                            If intN < gSDMax Then
                                                .Selection.TypeText("NA")
                                            Else
                                                If boolLUseSigFigs Then
                                                    .Selection.TypeText(Text:=CStr(DisplayNum(SigFigOrDec(numSD, LSigFig, False), LSigFig, False)))
                                                Else
                                                    .Selection.TypeText(Text:=CStr(Format(numSD, GetRegrDecStr(LSigFig))))
                                                End If


                                            End If

                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If

                                    varNom = CDec(arrBCQCs(2, int12))

                                    If boolSTATSCV Then
                                        Try
                                            'record %CV
                                            int8 = int8 + 1
                                            If intN < gSDMax Then
                                                .Selection.TypeText("NA")
                                            Else
                                                numPrec = CalcCVPercent(numSD, numMean, intQCDec)
                                                .Selection.TypeText(Format(numPrec, strQCDec))


                                            End If

                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()
                                        Catch ex As Exception

                                        End Try

                                    End If
                                    If boolSTATSBIAS And boolSTATSMEAN Then
                                        Try
                                            'record %Bias
                                            int8 = int8 + 1
                                            'var5 = ((numMean / CDec(arrBCQCs(2, int12))) - 1) * 100
                                            '.Selection.TypeText(Text:=Format(var5, strQCDec))

                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            '.Selection.TypeText(Format(numBias, strQCDec))
                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                            '.Selection.TypeText(Text:=Format(RoundToDecimal(((numMean / CDec(arrBCQCs(2, Count3))) - 1) * 100, 1), strqcdec))
                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()


                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolTHEORETICAL And boolSTATSMEAN Then
                                        Try
                                            'record %Theor
                                            int8 = int8 + 1
                                            'var5 = CDec(Format(((numMean / CDec(arrBCQCs(2, int12))) - 1) * 100, strQCDec))
                                            '.Selection.TypeText(Text:=Format(100 + var5, strQCDec))
                                            '.Selection.TypeText(Text:=Format(RoundToDecimal(((numMean / CDec(arrBCQCs(2, Count3))) - 1) * 100, 1), strqcdec))

                                            numTheor = CalcREPercent(numMean, varNom, intQCDec)
                                            numTheor = 100 + CDec(numTheor)
                                            '.Selection.TypeText(Format(numTheor, strQCDec))
                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            .Selection.TypeText(Text:=NomConcZero(varNom, numTheor))

                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()
                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolSTATSDIFF And boolSTATSMEAN Then
                                        Try
                                            'record %Bias
                                            int8 = int8 + 1
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            .Selection.TypeText(Format(numBias, strQCDec))
                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()


                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If BOOLSTATSRE And boolSTATSMEAN Then
                                        Try
                                            'record %RE
                                            int8 = int8 + 1
                                            numBias = CalcREPercent(numMean, varNom, intQCDec)
                                            '.Selection.TypeText(Format(numBias, strQCDec))
                                            '20180430 LEE:
                                            'must account for Endogenous Cmpds, NomConc = 0
                                            .Selection.TypeText(Text:=NomConcZero(varNom, numBias))

                                            '.Selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
                                            .Selection.Tables.Item(1).Cell(int1 + int8, Count3 + 1).Select()


                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If boolSTATSN Then
                                        Try
                                            'record n
                                            int8 = int8 + 1
                                            .Selection.TypeText(Text:=CStr(intN))


                                        Catch ex As Exception

                                        End Try
                                    End If

                                    If Count3 >= ctQCs * int11 Then
                                    Else
                                        '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)
                                        '.Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, int8 - 1)
                                        '.Selection.Tables.Item(1).Cell(int1, 1).Select()
                                        '.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1)
                                    End If

                                Next
                            End If
                        End If

                        'Call DeleteTableRows(wd)

                        '20180523 LEE
                        Call RemoveRows(wd, ctTbl)

                    Catch ex As Exception

                        str1 = "There was a problem preparing table:"
                        str1 = strM1 & ChrW(10) & ChrW(10) & str1
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                        MsgBox(str1, vbInformation, "Problem...")

                    End Try


                    'border bottom of thellos table
                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow)
                    ''
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                    .Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

                    'If boolQCREPORTACCVALUES Then
                    'Else
                    '    If intExp = 0 Then
                    '        Call DeleteRows(ctExp, wd)
                    '    End If
                    'End If

                    'enter table number
                    str1 = "Summary of " & arrAnalytes(14, Count1) & " Interpolated QC Standard Concentrations"

                    'Dim Oldrng as Microsoft.Office.Interop.Word.Range
                    'Oldrng = wd.Selection


                    '***
                    'strA = arrAnalytes(14, Count1)
                    If gNumMatrix = 1 Then
                        strA = strAnalC
                    Else
                        strA = strAnal 'strAnalC has '..Matrix', don't want to pass that here
                    End If
                    'No. Now just send strAnal
                    strA = strAnal
                    strTName = UpdateAnalyteMatrix(strTName, strA, strMatrix, True, intGroup, False)
                    Call EnterTableNumber(wd, strTName, 5, strA, strTempInfo, intTableID, intGroup, idTR)
                    'Note: strTName is byRef and will return Table, number, caption, label

                    '***

                    'return to old spot
                    'Oldrng.Select()

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

                    'split table, if needed
                    str1 = frmH.lblProgress.Text

                    ctLegend = ctLegend + 1
                    intLeg = intLeg + 1
                    arrLegend(1, intLeg) = "NA"
                    arrLegend(2, intLeg) = "Not Applicable"
                    arrLegend(3, intLeg) = False
                    arrLegend(4, intLeg) = False

                    'autofit table
                    Call AutoFitTable(wd, BOOLINCLUDEDATE)

                    strM = "Finalizing " & strTName & "..."
                    strM1 = strM & ChrW(10) & "Table " & intTCur & " of " & intTTot & " tables..."
                    str1 = strM1

                    frmH.lblProgress.Text = strM1
                    frmH.Refresh()

                    '''''''''''''wdd.visible = True


                    Call SplitTable(wd, 4, intLeg, arrLegend, str1, False, intLeg + 2, False, False, False, intTableID)
                    'Sub SplitTable(ByVal wd As Word.Application, ByVal ctHdRows As Short, ByVal ctLegend As Short, 
                    'ByVal arr As Object, ByVal strT As String, ByVal DoLegend As Boolean, ByVal intSplitRows As Short, ByVal boolSmallFont As Boolean)

                    ''''''''''wdd.visible = True

                    'autofit table
                    Call AutoFitTable(wd, False)

                    Call MoveOneCellDown(wd)

                    Call InsertLegend(wd, intTableID, idTR, False, 1)


                End If
end1:
next1:

skip4a:
            Next Count1


skip4:

        End With



    End Sub

End Module
