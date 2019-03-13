Option Compare Text

Imports System
Imports System.IO
Imports System.IO.FileSystemInfo
Imports Word = Microsoft.Office.Interop.Word
Imports System.Text
Imports System.Text.RegularExpressions

Module modFunctions

    Function GetDilQCInfo(ByVal intType As Short, ByVal intGroup As Short, ByVal idTR As Int32) As String

        '20190228 LEE:
        'intType: 1 = # of Diln QC Replicates
        'intType: 2 = Diln QC Concentration(s)
        'intType: 3 = Diln QC Dilution Factor

        Dim tbl1 As DataTable
        Dim tbl2 As DataTable
        Dim intT As Short
        Dim strF As String
        Dim strS As String
        Dim int1 As Int16
        Dim int2 As Int16
        Dim var1, var2, var3, var4
        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim rows1() As DataRow
        Dim intMax As Short = 0

        tbl1 = tblAssignedSamples
        intT = 12 '12 = Diln table
        strS = "ID_TBLCONFIGREPORTTABLES ASC"

        ' ''20190222 LEE: Need to find Dilution table in tblTableProperties
        'strF = GetBOOLSTATSNRFilter(intT)
        'strF = "INTGROUP = " & intGroup & " AND " & strF
        strF = "INTGROUP = " & intGroup & " AND ID_TBLREPORTTABLE = " & idTR

        Select Case intType
            Case 1 '# of Diln QC Replicates


                strS = "ID_TBLCONFIGREPORTTABLES ASC"
                Try
                    Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                    tbl2 = dv1.ToTable("a", True, "NOMCONC", "RUNID", "CHARANALYTE")
                    int1 = tbl2.Rows.Count
                    If int1 = 0 Then
                        GetDilQCInfo = "[NA]"
                    Else
                        'loop to find largest value

                        For Count2 = 0 To tbl2.Rows.Count - 1
                            var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                            var2 = tbl2.Rows.Item(0).Item("RUNID")
                            var3 = tbl2.Rows.Item(0).Item("CHARANALYTE")
                            strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                            strF = "ID_TBLREPORTTABLE = " & idTR & " AND NOMCONC = " & var1 & " AND RUNID = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                            rows1 = tbl1.Select(strF)
                            int2 = rows1.Length
                            If int2 > intMax Then
                                intMax = int2
                            End If
                        Next

                        If intMax = 0 Then
                            GetDilQCInfo = "[NA]"
                        Else
                            GetDilQCInfo = intMax ' VerboseNumber(intmax, False)
                        End If

                    End If
                Catch ex As Exception

                End Try

            Case 2

                'strF = "(ID_TBLCONFIGREPORTTABLES = " & intT & " OR ID_TBLCONFIGREPORTTABLES = 31) AND ID_TBLSTUDIES = " & idTS
                strS = "NOMCONC ASC"
                Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                tbl2 = dv1.ToTable("a", True, "NOMCONC", "CHARANALYTE")
                int1 = tbl2.Rows.Count
                If int1 = 0 Then
                    GetDilQCInfo = "[NA]"
                Else
                    'loop to find all values
                    For Count2 = 0 To tbl2.Rows.Count - 1
                        var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                        If Count2 = 0 Then
                            GetDilQCInfo = var1
                        ElseIf Count2 = tbl2.Rows.Count - 1 Then
                            GetDilQCInfo = GetDilQCInfo & " and " & var1
                        Else
                            GetDilQCInfo = GetDilQCInfo & ", " & var1
                        End If
                    Next
                    var1 = tbl2.Rows.Item(0).Item("NOMCONC")
                    GetDilQCInfo = var1
                End If

            Case 3

                strS = "ALIQUOTFACTOR DESC"
                Dim rowsD() As DataRow = tbl1.Select(strF, strS)
                Dim dtblT As DataTable = rowsD.CopyToDataTable
                Dim dtblD As DataTable = dtblT.DefaultView.ToTable("a", True, "ALIQUOTFACTOR")
                'Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
                Dim dv1 As System.Data.DataView = New DataView(dtblD, "", strS, DataViewRowState.CurrentRows)
                If dv1.Count = 0 Then
                    GetDilQCInfo = "[NA]"
                Else
                    Dim intDC As Short = 0
                    For Count2 = 0 To dv1.Count - 1
                        var1 = dv1(Count2).Item("ALIQUOTFACTOR")
                        'var2 = CInt(1 / var1)
                        var2 = GetDilnFactor(CDec(var1))
                        If dv1.Count > 2 Then
                            If Count2 = 0 Then
                                GetDilQCInfo = var2
                            ElseIf Count2 = dv1.Count - 1 Then
                                GetDilQCInfo = GetDilQCInfo & ", and " & var2
                            Else
                                GetDilQCInfo = GetDilQCInfo & ", " & var2
                            End If
                        Else
                            If Count2 = 0 Then
                                GetDilQCInfo = var2
                            Else
                                GetDilQCInfo = GetDilQCInfo & " and " & var2
                            End If
                        End If

                    Next

                End If

        End Select

    End Function


    Function GetDilnFactor(ByVal numX As Decimal)

        Try
            GetDilnFactor = Format(1 / numX, "0")
        Catch ex As Exception
            GetDilnFactor = 0
        End Try

    End Function

    Function GetBOOLSTATSNRFilter(ByVal intBS As Short) As String

        '20190225 LEE: Will return a filter string for appropriate ID_TBLREPORTTABLE for the provided BOOLSTATSNR value
        Dim strF As String
        Dim int1 As Int16
        Dim var1

        'intBS
        '1    rbNA    -1 or 0 or 1   
        '2    rbProcess    Extract (Process)   CHARPROCSTABILITY
        '3    rbBenchTop    BenchTop   CHARSTABILITYUNDERSTORAGECOND
        '4    rbFT    FreezeThaw   CHARDEMONSTRATEDFREEZETHAW
        '5    rbLT    LongTerm   CHARLTSTORSTAB
        '6    rbReinjection    Reinjection   CHARREFRSTAB
        '7    rbBlood    Blood   CHARBLOOD
        '8    rbStockSolution    StockSolution   CHARSTOCKSOLUTION
        '9    rbSpiking    Spiking   CHARSPIKING
        '10    rbAutosampler    Autosampler   CHARAUTOSAMPLER
        '11    rbBatchReinjection    Batch Reinjection   CHARBATCHREINJECTION
        '12    rbDilution       Dilution Samples    'no stability conditions

        ''20190222 LEE: Need to find Dilution table in tblTableProperties
        Dim dtblTP As DataTable = tblTableProperties
        Dim rowsTP() As DataRow
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLSTATSNR = " & intBS
        rowsTP = dtblTP.Select(strF)
        int1 = rowsTP.Length
        If int1 = 0 Then
            strF = "ID_TBLREPORTTABLE = 0"
        Else
            var1 = rowsTP(0).Item("ID_TBLREPORTTABLE")
            strF = "ID_TBLREPORTTABLE = " & var1
        End If

        GetBOOLSTATSNRFilter = strF

    End Function

    Function boolHasSTATSNR(ByVal intBS As Short, ByVal idCR As Int16, ByVal idTR As Int16) As Boolean

        boolHasSTATSNR = False

        '20190225 LEE: Will return a boolean if an ID_TBLREPORTTABLE for the provided BOOLSTATSNR value exists
        Dim strF As String
        Dim int1 As Int16
        Dim var1

        'intBS
        '1    rbNA    -1 or 0 or 1   
        '2    rbProcess    Extract (Process)   CHARPROCSTABILITY
        '3    rbBenchTop    BenchTop   CHARSTABILITYUNDERSTORAGECOND
        '4    rbFT    FreezeThaw   CHARDEMONSTRATEDFREEZETHAW
        '5    rbLT    LongTerm   CHARLTSTORSTAB
        '6    rbReinjection    Reinjection   CHARREFRSTAB
        '7    rbBlood    Blood   CHARBLOOD
        '8    rbStockSolution    StockSolution   CHARSTOCKSOLUTION
        '9    rbSpiking    Spiking   CHARSPIKING
        '10    rbAutosampler    Autosampler   CHARAUTOSAMPLER
        '11    rbBatchReinjection    Batch Reinjection   CHARBATCHREINJECTION
        '12    rbDilution       Dilution Samples    'no stability conditions

        ''20190222 LEE: Need to find Dilution table in tblTableProperties
        Dim dtblTP As DataTable = tblTableProperties
        Dim rowsTP() As DataRow
        'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLSTATSNR = " & intBS & " AND ID_TBLCONFIGREPORTTABLES = " & idCR
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLSTATSNR = " & intBS & " AND ID_TBLREPORTTABLE = " & idTR
        rowsTP = dtblTP.Select(strF)
        int1 = rowsTP.Length
        If int1 = 0 Then
            boolHasSTATSNR = False
        Else
            boolHasSTATSNR = True
        End If

    End Function


    Function GetStabPeriod(ByVal dtbl As System.Data.DataTable, ByVal idS As Int64, ByVal idT As Int64, ByVal strCol As String) As String

        GetStabPeriod = ""

        '20181111 LEE:
        'New logic
        'tblTableProperties.boolStatsNR contains stability experiment type information
        'tblTableProperties.CHARCARRYOVERLABEL has notes
        'tblMethodValidationData.
        'idT will be ignored

        'legend
        'Case = str1/strCol
        '                        Case "Freeze/Thaw Stability in Matrix"
        '                            str2 = GetStabPeriod(dtblT, id_tblStudies, 19, str1)
        '                            'remove any parenthases
        '                            str2 = Replace(str2, "(", "")
        '                            str3 = Replace(str2, ")", "")
        '                            strV = str3
        '                        Case "Maximum # of Freeze/thaw Cycles"
        '
        '                        Case "Stability under Storage Conditions" 'deprecated, is now benchtop stability 20181110
        '                        Case "Bench-top Stability"
        '                            strV = GetStabPeriod(dtblT, id_tblStudies, 21, str1)
        '                        Case "Is Stability >= Maximum Storage Duration"
        '                        Case "Process Stability"
        '                            strV = GetStabPeriod(dtblT, id_tblStudies, 21, str1)
        '                        Case "Refrigerated Stability in Matrix" 'deprecated now Reinjection Stability 20181110
        '                        Case "Reinjection Stability"
        '                            strV = GetStabPeriod(dtblT, id_tblStudies, 18, str1)
        '                        Case "Long-term Storage Stability in Matrix"
        '                            strV = GetStabPeriod(dtblT, id_tblStudies, 29, str1)

        'legend
        '1 rbNA  boolNA -1 or 0 or 1
        '2 rbProcess  boolProcess
        '3 rbBenchtop  boolBenchtop
        '4 rbFT  boolFT
        '5 rbLT  boolLT
        '6 rbReinjection  boolReinjection
        '7 rbBlood  boolBlood
        '8 rbStockSolution  boolStockSolution
        '9 rbSpiking  boolSpiking

        '20190109 LEE:
        '1    rbNA    -1 or 0 or 1   
        '2    rbProcess    Extract (Process)   CHARPROCSTABILITY
        '3    rbBenchTop    BenchTop   CHARSTABILITYUNDERSTORAGECOND
        '4    rbFT    FreezeThaw   CHARDEMONSTRATEDFREEZETHAW
        '5    rbLT    LongTerm   CHARLTSTORSTAB
        '6    rbReinjection    Reinjection   CHARREFRSTAB
        '7    rbBlood    Blood   CHARBLOOD
        '8    rbStockSolution    StockSolution   CHARSTOCKSOLUTION
        '9    rbSpiking    Spiking   CHARSPIKING
        '10    rbAutosampler    Autosampler   CHARAUTOSAMPLER
        '11    rbBatchReinjection    Batch Reinjection   CHARBATCHREINJECTION
        '12    rbDilution       Dilution Samples    'no stability conditions



        'field code legend: e.g.[CHARDEMONSTRATEDFREEZETHAW]
        '                        Case "Freeze/Thaw Stability in Matrix"
        '                            strV = rowsD(0).Item("CHARDEMONSTRATEDFREEZETHAW")
        '                        Case "Bench-top Stability"
        '                            strV = rowsD(0).Item("CHARSTABILITYUNDERSTORAGECOND")
        '                        Case "Process Stability"
        '                            strV = rowsD(0).Item("CHARPROCSTABILITY")
        '                        Case "Reinjection Stability"
        '                            strV = rowsD(0).Item("CHARREFRSTAB")
        '                        Case "Long-term Storage Stability in Matrix"
        '                            strV = rowsD(0).Item("CHARLTSTORSTAB")


        'this is the logic:
        'in Advanced Table Configuration, Storage Conditions statement is stored in tblReportTable.CHARSTABILITYPERIOD
        'This value will be entered into tblMethodValidation.[xxx] if [xxx] is empty
        'User can overwrite tblMethodValidation.[xxx] in the Review Method Validation window
        'If user wants to put CHARSTABILITYPERIOD back, then:
        '  - delete contents in Review Method Validation window
        '  - Re-load the study

        '20190215 LEE: if boolProp = true, then get from strField

        Dim intR As Short
        Dim strFD As String
        Dim idCRT As Short = 0
        intR = 0
        idT = 0

        Select Case strCol
            Case "Freeze/Thaw Stability"
                intR = 4
                strFD = "CHARDEMONSTRATEDFREEZETHAW"

            Case "Maximum # of Freeze/thaw Cycles"
                intR = -1
                strFD = "INTNUMBEROFCYCLES"
                idCRT = 19
            Case "Bench-top Stability"
                intR = 3
                strFD = "CHARSTABILITYUNDERSTORAGECOND"
            Case "Process Stability"
                intR = 2
                strFD = "CHARPROCSTABILITY"
            Case "Reinjection Stability"
                intR = 6
                strFD = "CHARREFRSTAB"
            Case "Long-term Storage Stability"
                intR = 5
                strFD = "CHARLTSTORSTAB"

                '20190109 LEE
            Case "Whole Blood Stability"
                intR = 7
                strFD = "CHARBLOOD"
            Case "Stock Solution Stability"
                intR = 8
                strFD = "CHARSTOCKSOLUTION"
            Case "Spiking Solution Stability"
                intR = 9
                strFD = "CHARSPIKING"
            Case "Autosampler Stability"
                intR = 10
                strFD = "CHARAUTOSAMPLER"
            Case "Batch Reinjection Stability"
                intR = 11
                strFD = "CHARBATCHREINJECTION"

        End Select


        If intR = 0 Then
            GoTo end1
        End If


        Dim strF As String
        Dim strF1 As String
        Dim rows() As DataRow
        Dim intRows As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim int1 As Int16
        Dim int2 As Int16
        Dim int3 As Int16
        Dim strT As String = ""
        Dim var1

        Dim dtblTP As DataTable = tblTableProperties
        Dim rowsTP() As DataRow
        Dim dtblRT As DataTable = tblReportTable

        '20190110 LEE:
        'Aack! Logic must be changed
        'Current logic is returning data for unincluded tables
        'Must evaluate tblReportTable

        int1 = 0

        If intR = -1 Then
            '20190215 LEE:
            'must get info from tblTableProperties
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCRT
            rowsTP = dtblTP.Select(strF)

            For Count1 = 1 To rowsTP.Length

                'check to ensure table is included
                idT = rowsTP(Count1 - 1).Item("ID_TBLREPORTTABLE")
                Dim rowsRT() As DataRow = dtblRT.Select("ID_TBLREPORTTABLE = " & idT & " AND BOOLINCLUDE = -1")
                If rowsRT.Length > 0 Then
                    var1 = NZ(rowsTP(Count1 - 1).Item(strFD), "")
                    If Len(var1) = 0 Then
                    Else
                        int1 = int1 + 1
                        If int1 = 1 Then
                            strT = CStr(var1)
                        Else
                            strT = strT & ChrW(13) & CStr(var1)
                        End If
                    End If
                End If

            Next Count1

        Else

            If intR = -1 Then
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 19"
            Else
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLSTATSNR = " & intR
            End If

            rowsTP = dtblTP.Select(strF)
            strT = ""
            int1 = 0
            For Count1 = 1 To rowsTP.Length

                'get CHARSTABILITYPERIOD from tblReportTables
                idT = rowsTP(Count1 - 1).Item("ID_TBLREPORTTABLE")
                Dim rowsRT() As DataRow = dtblRT.Select("ID_TBLREPORTTABLE = " & idT & " AND BOOLINCLUDE = -1")
                If rowsRT.Length = 0 Then
                Else
                    strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
                    rows = tblReportTable.Select(strF1)
                    If rows.Length = 0 Then
                    Else
                        str1 = NZ(rows(0).Item("CHARSTABILITYPERIOD"), "")
                        If Len(str1) = 0 Then
                        Else
                            int1 = int1 + 1
                            If int1 = 1 Then
                                strT = str1
                            Else
                                'strT = strT & ChrW(10) & str1
                                strT = strT & ChrW(13) & str1
                            End If
                        End If

                    End If

                End If

            Next Count1
        End If





        GetStabPeriod = strT


end1:

    End Function

    Function GetStatsNR(idRT As Int32)

        GetStatsNR = 1

        Dim strF As String

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idRT
        Dim rows() As DataRow = tblTableProperties.Select(strF)

        If rows.Length = 0 Then
            GetStatsNR = 1
        Else
            GetStatsNR = NZ(rows(0).Item("BOOLSTATSNR"), 1)
        End If

    End Function


    Function ReturnMForME() As String

        '20181108 LEE:

        If BOOLINCLINTSTDNMF Then
            ReturnMForME = "Matrix Effect"
        Else
            ReturnMForME = "Matrix Factor"
        End If

    End Function

    Function ReturnMForMEAbbr() As String

        '20181108 LEE:

        If BOOLINCLINTSTDNMF Then
            ReturnMForMEAbbr = "ME"
        Else
            ReturnMForMEAbbr = "MF"
        End If

    End Function

    Function ReturnQCGroup(boolA, boolE, boolR, a, b, c, d, e, f) As String

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

        Dim strF As String = ""

        If boolA Then

            If boolE Then
                If INTQCLEVELGROUP = 0 Then 'use assaylevel
                    strF = "ASSAYLEVEL = " & b & " AND ALIQUOTFACTOR = " & NZ(c, 1) & " AND ELIMINATEDFLAG = '" & f & "'"
                ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                    strF = "NOMCONC = " & d & " AND ALIQUOTFACTOR = " & NZ(c, 1) & " AND ELIMINATEDFLAG = '" & f & "'"
                ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                    strF = "CHARHELPER1 = '" & e & "' AND ALIQUOTFACTOR = " & NZ(c, 1) & " AND ELIMINATEDFLAG = '" & f & "'"
                Else
                    strF = "ASSAYLEVEL = " & b & " AND ALIQUOTFACTOR = " & NZ(c, 1) & " AND ELIMINATEDFLAG = '" & f & "'"
                End If
            Else
                If boolR Then
                    If INTQCLEVELGROUP = 0 Then 'use assaylevel
                        strF = " AND RUNID = " & a & " AND ASSAYLEVEL = " & b & " AND ALIQUOTFACTOR = " & NZ(c, 1)
                    ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                        strF = " AND RUNID = " & a & " AND NOMCONC = " & d & " AND ALIQUOTFACTOR = " & NZ(c, 1)
                    ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                        strF = " AND RUNID = " & a & " AND CHARHELPER1 = '" & e & "' AND ALIQUOTFACTOR = " & NZ(c, 1)
                    Else
                        strF = " AND RUNID = " & a & " AND ASSAYLEVEL = " & b & " AND ALIQUOTFACTOR = " & NZ(c, 1)
                    End If
                Else
                    If INTQCLEVELGROUP = 0 Then 'use assaylevel
                        strF = "ASSAYLEVEL = " & b & " AND ALIQUOTFACTOR = " & NZ(c, 1)
                        'strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & int20 & " AND ASSAYLEVEL = " & var2 & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS <> 4"
                    ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                        strF = "NOMCONC = " & d & " AND ALIQUOTFACTOR = " & NZ(c, 1)
                    ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                        strF = "CHARHELPER1 = '" & e & "' AND ALIQUOTFACTOR = " & NZ(c, 1)
                    Else
                        strF = "ASSAYLEVEL = " & b & " AND ALIQUOTFACTOR = " & NZ(c, 1)
                    End If
                End If

            End If

        Else

            If INTQCLEVELGROUP = 0 Then 'use assaylevel
                strF = " AND RUNID = " & a & " AND ASSAYLEVEL = " & b
            ElseIf INTQCLEVELGROUP = 1 Then 'use NomConc
                '20171118 LEE: need aliquot factor too because sometimes NomConc is same for Diln and Hi samples
                strF = " AND RUNID = " & a & " AND NOMCONC = " & d & " AND ALIQUOTFACTOR = " & NZ(c, 1)
            ElseIf INTQCLEVELGROUP = 2 Then 'use Level Label
                strF = " AND RUNID = " & a & " AND CHARHELPER1 = '" & e & "'"
            Else
                strF = " AND RUNID = " & a & " AND ASSAYLEVEL = " & b
            End If
        End If

        ReturnQCGroup = strF

    End Function

    Function ReturnFCID(rng As Word.Range) As String

        ReturnFCID = ""

        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim var1, var2, var3, var4
        Dim strFind As String
        Dim c As Word.Range
        Dim idRT As Int32


        strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND CHARFCID IS NOT NULL"

        Dim rowsRT() As DataRow = tblReportTable.Select(strF1)
        Dim rowsRTA() As DataRow

        Dim Count1 As Int16
        Dim Count2 As Int16

        For Count1 = 0 To rowsRT.Length - 1
            var1 = NZ(rowsRT(Count1).Item("CHARFCID"), "")
            idRT = rowsRT(Count1).Item("ID_TBLREPORTTABLE")
            If Len(var1) = 0 Then
            Else
                strFind = "_" & var1 & "]"
                'look for strfind in rng
                With rng.Find
                    .ClearFormatting()
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .Format = True
                    .MatchCase = True
                    .Execute(FindText:=strFind)
                    If .Found Then
                        ReturnFCID = var1
                        .ClearFormatting()
                        Exit For
                    End If
                    .ClearFormatting()

                End With

            End If
        Next Count1



    End Function

    Function IsAnalInTable(rng As Word.Range, INTGROUP As Short, strAnal As String) As Boolean

        IsAnalInTable = False

        '20180815 LEE:

        'find available field codes
        'tblReportTable has CHARFCID and BOOLINCLUDE (included in the study) and ID_TBLREPORTTABLE
        'tblReportTableAnalytes has ID_TBLREPORTTABLE and BOOLINLCUDE (analyte included in the table) and INTGROUP and ANALYTEID

        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim var1, var2, var3, var4
        Dim strFind As String
        Dim idRT As Int32
        Dim str1 As String

        strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND CHARFCID IS NOT NULL"

        Dim rowsRT() As DataRow = tblReportTable.Select(strF1)
        Dim rowsRTA() As DataRow

        Dim Count1 As Int16
        Dim Count2 As Int16

        Dim boolHit As Boolean = False

        str1 = rng.Text

        For Count1 = 0 To rowsRT.Length - 1
            var1 = NZ(rowsRT(Count1).Item("CHARFCID"), "")
            idRT = rowsRT(Count1).Item("ID_TBLREPORTTABLE")
            If Len(var1) = 0 Then
            Else
                strFind = "_" & var1 & "]"
                'look for strfind in rng
                If InStr(1, str1, strFind, CompareMethod.Text) > 0 Then
                    boolHit = True
                    'now check to see if this analyte is included for the table
                    strF2 = "ID_TBLREPORTTABLE = " & idRT & " AND INTGROUP = " & INTGROUP & " AND BOOLINCLUDE = -1"
                    rowsRTA = tblReportTableAnalytes.Select(strF2)
                    If rowsRTA.Length = 0 Then
                    Else
                        IsAnalInTable = True
                        Exit For
                    End If
                End If

            End If

        Next Count1


        If IsAnalInTable Then
        Else
            If boolHit Then
            Else
                'then is just ANALYTE_x
                IsAnalInTable = True
            End If
        End If

        'For Count1 = 0 To rowsRT.Length - 1
        '    var1 = NZ(rowsRT(Count1).Item("CHARFCID"), "")
        '    idRT = rowsRT(Count1).Item("ID_TBLREPORTTABLE")
        '    If Len(var1) = 0 Then
        '    Else
        '        strFind = "_" & var1 & "]"
        '        'look for strfind in rng
        '        With rng.Find
        '            .ClearFormatting()
        '            .Forward = True
        '            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
        '            .Format = True
        '            .MatchCase = True
        '            .Execute(FindText:=strFind)
        '            If .Found Then
        '                boolHit = True
        '                'now check to see if this analyte is included for the table
        '                strF2 = "ID_TBLREPORTTABLE = " & idRT & " AND INTGROUP = " & INTGROUP & " AND BOOLINCLUDE = -1"
        '                rowsRTA = tblReportTableAnalytes.Select(strF2)
        '                If rowsRTA.Length = 0 Then
        '                Else
        '                    IsAnalInTable = True
        '                    .ClearFormatting()
        '                    Exit For
        '                End If
        '            End If
        '            .ClearFormatting()

        '        End With

        '    End If
        'Next Count1

        'If IsAnalInTable Then
        'Else
        '    If boolHit Then
        '    Else
        '        'then is just ANALYTE_x
        '        IsAnalInTable = True
        '    End If
        'End If


    End Function

    Function ReturnDiff(ByVal numA2 As Decimal, ByVal numB2 As Decimal) As Decimal

        '20180803 LEE:

        Dim var1, var2, var3

        Dim boolOld As Boolean = BOOLCALCINTSTDNMF 'This is denominator
        'rbOld: BOOLCALCINTSTDNMF  = true
        'rbNew: BOOLCALCINTSTDNMF  = false

        If boolMEANACCURACY Then
            If numA2 And numB2 = 0 Then
                var3 = 0
            Else

                If boolPOSLEG Then
                    var3 = (numA2 - numB2) / ((numA2 + numB2) / 2) * 100
                Else
                    var3 = (numB2 - numA2) / ((numA2 + numB2) / 2) * 100
                End If
            End If

        ElseIf BOOLDIFFERENCE Then

            If boolOld Then
                If numA2 = 0 Then
                    var3 = 0
                Else
                    If boolPOSLEG Then
                        var3 = RoundToDecimalRAFZ((numA2 - numB2) / numA2 * 100, 10) '(T0-Tn)/((Tn+T0)/2)
                    Else
                        var3 = RoundToDecimalRAFZ((numB2 - numA2) / numA2 * 100, 10) '(Tn-T0)/((Tn+T0)/2)
                    End If
                End If
            Else
                If numB2 = 0 Then
                    var3 = 0
                Else
                    If boolPOSLEG Then
                        var3 = RoundToDecimalRAFZ((numA2 - numB2) / numB2 * 100, 10) '(T0-Tn)/((Tn+T0)/2)
                    Else
                        var3 = RoundToDecimalRAFZ((numB2 - numA2) / numB2 * 100, 10) '(Tn-T0)/((Tn+T0)/2)
                    End If
                End If
            End If

        ElseIf boolRECOVERY Then

            If boolOld Then
                If boolPOSLEG Then
                    If numA2 = 0 Then
                        var3 = 0
                    Else
                        var3 = RoundToDecimalRAFZ(numA2 / numA2 * 100, 10) 'T0/Tn
                    End If
                Else
                    If numA2 = 0 Then
                        var3 = 0
                    Else
                        var3 = RoundToDecimalRAFZ(numB2 / numA2 * 100, 10) 'Tn/T0
                    End If
                End If
            Else
                If boolPOSLEG Then
                    If numB2 = 0 Then
                        var3 = 0
                    Else
                        var3 = RoundToDecimalRAFZ(numA2 / numB2 * 100, 10) 'T0/Tn
                    End If
                Else
                    If numB2 = 0 Then
                        var3 = 0
                    Else
                        var3 = RoundToDecimalRAFZ(numB2 / numB2 * 100, 10) 'Tn/T0
                    End If
                End If
            End If

        Else


        End If

        ReturnDiff = CDec(var3)


    End Function

    Function StringRepDegC(strX As String) As String

        StringRepDegC = Replace(strX, "degC", ChrW(176) & "C", 1, -1, CompareMethod.Text)
        StringRepDegC = Replace(StringRepDegC, "deg C", ChrW(176) & fNBSP() & "C", 1, -1, CompareMethod.Text)


    End Function

    Function NomConcZero(varNom, x) As String

        If NZ(varNom, 0) <= 0 Then
            NomConcZero = "NA"
        Else
            Try
                If IsNumeric(x) Then
                    NomConcZero = Format(CDec(x), strQCDec)
                Else
                    NomConcZero = "NA"
                End If
            Catch ex As Exception
                NomConcZero = "NA"
            End Try
        End If

    End Function

    Function GetAnalysisDateLabel(idCRT As Int64) As String

        '20180420 LEE:
        'User can now configure the text of (Analysis Date)

        GetAnalysisDateLabel = "Analysis Date"

        Dim idCHL As Int64

        Dim strF As String = "ID_TBLCONFIGREPORTTABLES = " & idCRT & " AND CHARCOLUMNLABEL = 'Analysis Date'"
        Dim rows() As DataRow = tblConfigHeaderLookup.Select(strF, "", DataViewRowState.CurrentRows)

        If rows.Length = 0 Then
        Else

            idCHL = rows(0).Item("ID_TBLCONFIGHEADERLOOKUP")

            Dim strF1 As String = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCRT & " AND ID_TBLCONFIGHEADERLOOKUP = " & idCHL

            Dim rows1() As DataRow = tblReportTableHeaderConfig.Select(strF1, "", DataViewRowState.CurrentRows)

            If rows1.Length = 0 Then
            Else
                GetAnalysisDateLabel = NZ(rows1(0).Item("CHARUSERLABEL"), "Analysis Date")
            End If
        End If



    End Function

    'Function GetGroup(intAnalyteID As Int32, strMatrix As String) As Short

    '    '20180628 LEE:
    '    'Aack! This doesn't account for different concentrations!!!

    '    GetGroup = -2

    '    Dim strF As String
    '    Dim tbl As DataTable = tblAnalyteGroups
    '    Dim var1, var2, var3
    '    Dim Count1 As Short

    '    strF = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
    '    Dim rows() As DataRow = tbl.Select(strF)

    '    If rows.Length = 0 Then
    '        GetGroup = -2
    '    ElseIf rows.Length = 1 Then
    '        var1 = rows(0).Item("INTGROUP")
    '        GetGroup = var1
    '    Else
    '        For Count1 = 0 To rows.Length - 1
    '            var1 = rows(Count1).Item("INTGROUP")
    '            var2 = rows(Count1).Item("CALIBRSET")
    '            If StrComp(var2, "-1(-1)", CompareMethod.Text) = 0 Then
    '            Else
    '                GetGroup = var1
    '                Exit For
    '            End If

    '        Next
    '    End If

    'End Function

    Function GetAN(x) As String

        Dim y As Decimal = CDec(x)

        GetAN = "a"
        Select Case y
            Case 8, 11
                GetAN = "an"
            Case x >= 80 Or x <= 89
                GetAN = "an"
            Case x >= 800 Or x <= 899
                GetAN = "an"
        End Select

    End Function

    Function ReturnPrecLabel()

        ReturnPrecLabel = "%CV"
        If BOOLUSERSD Then
            ReturnPrecLabel = "%RSD"
        End If

    End Function

    Function GetQCID(ByRef AnalyteID As Int64, ByRef intRunID As Int32, ByRef numConc As Single) As String

        GetQCID = ""

        Dim dtbl As DataTable = tblAllStdsAssay
        Dim strF As String
        Dim var1

        strF = "ANALYTEID = " & AnalyteID & " AND RUNID = " & intRunID & " AND CONCENTRATION = " & numConc
        Dim rows() As DataRow = dtbl.Select(strF)

        If rows.Length = 0 Then
        Else
            var1 = NZ(rows(0).Item("ID"), "")
            GetQCID = var1
        End If


    End Function

    Function AllowLockFinalReport() As Boolean

        AllowLockFinalReport = True

        If boolFormLoad Then
            Exit Function
        End If

        AllowLockFinalReport = False

        Dim strF As String = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow = tblPermissions.Select(strF)

        Dim intP As Short = NZ(rows(0).Item("BOOLLOCKFINALREPORT"), 0)

        Dim strM As String

        If intP = 0 Then

            Dim boolF As Boolean = boolFormLoad
            boolF = boolFormLoad
            boolFormLoad = True
            frmH.chkLockFinalReport.Checked = Not (frmH.chkLockFinalReport.Checked)
            boolFormLoad = boolF

            strM = "User belongs to a Permissions Group that is not allowed to modify the Final Report lock status."
            MsgBox(strM, vbInformation, "Invalid action...")
        Else
            AllowLockFinalReport = True
        End If


    End Function


    Function FindRowInDGV(strCol As String, varVal As Object, dgv As DataGridView) As Int32

        Try
            FindRowInDGV = dgv.CurrentRow.Index
        Catch ex As Exception
            FindRowInDGV = 0
        End Try


        Dim Count1 As Int32
        Dim var1

        For Count1 = 0 To dgv.Rows.Count - 1
            var1 = dgv(strCol, Count1).Value
            If var1 = varVal Then
                FindRowInDGV = Count1
                Exit For
            End If

        Next

    End Function

    Function EnterBackSlash(strX As String) As String

        EnterBackSlash = strX

        If Len(strX) = 0 Then
            GoTo end1
        End If

        Dim str1 As String
        str1 = Mid(strX, Len(strX), 1)
        If StrComp(str1, "\", CompareMethod.Text) = 0 Then
        Else
            EnterBackSlash = strX & "\"
        End If

end1:

    End Function

    Function GetChromReporter() As String

        GetChromReporter = ""

        Dim strM As String = ""
        Dim str1 As String

        Dim strF As String

        strF = "ID_TBLCONFIGURATION = 35"
        Dim rows() As DataRow = tblConfiguration.Select(strF)

        If rows.Length = 0 Then
            strM = "ChromReporter has not been configured in StudyDoc."
            GoTo end1
        End If

        str1 = NZ(rows(0).Item("CHARCONFIGVALUE"), "")

        If Len(str1) = 0 Then
            strM = "ChromReporter has not been configured in StudyDoc."
            GoTo end1
        End If

        'determine of chromreportpath exists
        If File.Exists(str1) Then
        Else
            strM = "ChromReporter has  been configured in StudyDoc, but the configure path to the ChromReport executable" & ChrW(10) & ChrW(10) & str1 & ChrW(10) & ChrW(10)
            strM = strM & "does not exist."
            GoTo end1
        End If

        GetChromReporter = str1

end1:

        If Len(strM) > 0 Then
            MsgBox(strM, vbInformation, "Invalid action...")
        End If


    End Function

    Function PasswordEncrypt(ByVal strP As String) As String

        PasswordEncrypt = ""

        Try
            PasswordEncrypt = Decode(Coding(strP, True), False)
        Catch ex As Exception

        End Try

    End Function

    Function PasswordUnEncrypt(ByVal strP As String) As String

        PasswordUnEncrypt = ""

        Try
            PasswordUnEncrypt = Coding(Decode(strP, True), False)
        Catch ex As Exception

        End Try

    End Function

    Function AreaRatioCalibr(intAnalyteID As Int64, intRunID As Int32) As Boolean

        AreaRatioCalibr = True

        Dim strF As String
        strF = "ANALYTEID = " & intAnalyteID & " AND RUNID = " & intRunID & "AND RUNSAMPLEKIND = 'STANDARD' AND INTERNALSTANDARDAREA > 0"

        Dim rows() As DataRow = tblAnalysisResultsHome.Select(strF)
        If rows.Length > 0 Then
        Else
            AreaRatioCalibr = False
        End If

    End Function

    Function GetFlagPercent(ASSAYID As Int32, ANALTYEID As Int32, LEVELNUMBER As Int16, numNomConc As Decimal, NUMRUNID As Int16) As Decimal

        GetFlagPercent = 15

        Dim var1, var2, var3, var4, var5

        Try

            Dim strFP As String
            strFP = "ASSAYID = " & ASSAYID & " AND ANALYTEID = " & ANALTYEID & " AND LEVELNUMBER = " & LEVELNUMBER

            Dim str1 As String
            Dim str2 As String

            Dim rowsAssID() As DataRow = tblBCStdsAssayIDAll.Select(strFP, "LEVELNUMBER ASC")
            Dim Count5 As Int16

            Dim boolRAID As Boolean = False
            Dim dec1 As Decimal

            For Count5 = 0 To rowsAssID.Length - 1
                dec1 = NZ(rowsAssID(Count5).Item("CONCENTRATION"), 0)
                If dec1 = numNomConc Then
                    boolRAID = True
                    Exit For
                End If
            Next
            If boolRAID Then
                GetFlagPercent = NZ(rowsAssID(Count5).Item("ANALYTEFLAGPERCENT"), 15) 'debugging
            Else
                'get from ASSAYREPS
                str1 = "ASSAYID = " & ASSAYID & " AND LEVELNUMBER = " & LEVELNUMBER
                Dim rowsAR() As DataRow = tblASSAYREPS.Select(str1)
                If rowsAR.Length = 0 Then
                    GetFlagPercent = rowsAR(0).Item("FLAGPERCENT")
                Else
                    GetFlagPercent = 15
                End If

            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

    End Function

    Function GetWatsonColH(ByVal idCT As Int32) As String

        GetWatsonColH = "Watson Run ID"

        Dim str1 As String
        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim id1 As Int64

        Try
            'first get Watson Run ID from tblConfigHeaderLookup
            strF1 = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND CHARCOLUMNLABEL = 'Watson Run ID'"
            Dim rowsA() As DataRow = tblConfigHeaderLookup.Select(strF1)

            If rowsA.Length = 0 Then
            Else

                id1 = rowsA(0).Item("ID_TBLCONFIGHEADERLOOKUP")
                'retrieve report table column header information
                strF = "id_tblStudies = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCT & " AND ID_TBLCONFIGHEADERLOOKUP = " & id1
                strS = "intOrder"
                Dim rows() As DataRow = tblReportTableHeaderConfig.Select(strF, strS)

                If rows.Length = 0 Then
                Else
                    GetWatsonColH = NZ(rows(0).Item("CHARUSERLABEL"), "Watson Run ID")
                End If

            End If
        Catch ex As Exception

        End Try


    End Function

    Function GetLabelColH(ByVal idCT As Int32) As String

        GetLabelColH = "Watson Run ID"

        Dim str1 As String
        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim id1 As Int64

        Try
            'first get Watson Run ID from tblConfigHeaderLookup
            strF1 = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND CHARCOLUMNLABEL = 'Label'"
            Dim rowsA() As DataRow = tblConfigHeaderLookup.Select(strF1)

            If rowsA.Length = 0 Then
            Else

                id1 = rowsA(0).Item("ID_TBLCONFIGHEADERLOOKUP")
                'retrieve report table column header information
                strF = "id_tblStudies = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCT & " AND ID_TBLCONFIGHEADERLOOKUP = " & id1
                strS = "intOrder"
                Dim rows() As DataRow = tblReportTableHeaderConfig.Select(strF, strS)

                If rows.Length = 0 Then
                Else
                    GetLabelColH = NZ(rows(0).Item("CHARUSERLABEL"), "")
                End If

            End If
        Catch ex As Exception

        End Try

    End Function



    Function GetSampleName(idCT As Int32) As String

        GetSampleName = ""

        Dim str1 As String
        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim id1 As Int64

        Try
            'first get Watson Run ID from tblConfigHeaderLookup
            strF1 = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND CHARCOLUMNLABEL = 'Sample Name'"
            Dim rowsA() As DataRow = tblConfigHeaderLookup.Select(strF1)

            If rowsA.Length = 0 Then
            Else

                id1 = rowsA(0).Item("ID_TBLCONFIGHEADERLOOKUP")
                'retrieve report table column header information
                strF = "id_tblStudies = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCT & " AND ID_TBLCONFIGHEADERLOOKUP = " & id1 & " AND BOOLINCLUDE = -1"
                strS = "intOrder"
                Dim rows() As DataRow = tblReportTableHeaderConfig.Select(strF, strS)

                If rows.Length = 0 Then
                Else
                    GetSampleName = NZ(rows(0).Item("CHARUSERLABEL"), "Sample Name")
                End If

            End If
        Catch ex As Exception

        End Try


    End Function

    Function GetAveDiffColAcc(ByVal numQCLevel As Short, ByVal dtblAccDiffCol As DataTable, ByVal intRunID As Int16, boolOutlier As Boolean) As Decimal

        GetAveDiffColAcc = 0

        Dim strFF As String
        'legend
        'Dim dtblAccDiffCol As New DataTable
        'Select Case Count1
        '    Case 1
        '        str1 = "numAcc"
        '    Case 2
        '        str1 = "boolOut"
        '    Case 3
        '        str1 = "QCLevel"
        '    Case 4
        '        str1 = "RunID"
        'End Select
        If boolOutlier Then
            If intRunID = 0 Then
                strFF = "QCLevel = " & numQCLevel
            Else
                strFF = "QCLevel = " & numQCLevel & " AND RunID = " & intRunID
            End If
        Else
            If intRunID = 0 Then
                strFF = "QCLevel = " & numQCLevel & " AND BOOLOUTLIER = FALSE"
            Else
                strFF = "QCLevel = " & numQCLevel & " AND RunID = " & intRunID & " AND BOOLOUTLIER = FALSE"
            End If
        End If

        Dim rowsFF() As DataRow = dtblAccDiffCol.Select(strFF)

        If rowsFF.Length = 0 Then
            GetAveDiffColAcc = -1
        Else

            Dim CountFF As Short
            Dim numFF As Decimal
            Dim numFFTot As Decimal = 0
            For CountFF = 0 To rowsFF.Length - 1
                numFF = rowsFF(CountFF).Item("numACC")
                numFFTot = numFFTot + numFF
            Next
            GetAveDiffColAcc = Format(CDec(numFFTot / rowsFF.Length), strQCDec)

        End If


    End Function

    Function GetBiasFromDiffCol(ByVal idTR As Int64, ByVal numNomConc As Single, ByVal int12 As Int16, NUMRUNID As Int16, boolInclOutlier As Boolean) As Decimal

        GetBiasFromDiffCol = 0

        'get numbias from average of %Bias columns
        'Call InsertQCTables(intTableID, idTR, charFCID, varNom, int12, "Accuracy", var3, CSng(numRunID), Count1, strDo, v1, v2, FALSE)
        'InsertQCTables(ByVal idCR As Int64, ByVal idT As Int64, ByVal charFCID As String, ByVal numNomConc As Single, ByVal numQCLevel As Single, 
        'ByVal charType As String, ByVal numValue As Single, ByVal numRunID As Single, ByVal numAnalyte As Int16, ByVal charAnalyte As String, 
        'ByVal numCrit1 As Single, numCrit2 As Single)
        ''legend
        'Dim row As DataRow = tbl.NewRow
        'id = id + 1
        'row.Item("ID_TBLQCTABLES") = id
        'row.Item("ID_TBLCONFIGREPORTTABLES") = idCR
        'row.Item("ID_TBLREPORTTABLE") = idT
        'row.Item("CHARFCID") = charFCID
        'row.Item("NUMNOMCONC") = numNomConc
        'row.Item("NUMQCLEVEL") = numQCLevel
        'row.Item("CHARTYPE") = charType
        'row.Item("numVALUE") = numValue
        'row.Item("NUMRUNID") = numRunID
        'row.Item("numAnalyte") = numAnalyte
        'row.Item("charAnalyte") = charAnalyte
        'row.Item("numCrit1") = numCrit1
        'row.Item("numCrit2") = numCrit1

        'tbl = tblQCTables

        Dim strF As String
        Dim var1, var2

        If boolInclOutlier Then
            strF = "ID_TBLREPORTTABLE = " & idTR & " AND NUMNOMCONC = " & numNomConc & " AND NUMQCLEVEL = " & int12 & " AND CHARTYPE = 'Accuracy'"
            If NUMRUNID = 0 Then
            Else
                strF = strF & " AND NUMRUNID = " & NUMRUNID
            End If
        Else
            strF = "ID_TBLREPORTTABLE = " & idTR & " AND NUMNOMCONC = " & numNomConc & " AND NUMQCLEVEL = " & int12 & " AND CHARTYPE = 'Accuracy' AND BOOLOUTLIER = FALSE"
            If NUMRUNID = 0 Then
            Else
                strF = strF & " AND NUMRUNID = " & NUMRUNID
            End If
        End If
        'strF = "CHARFCID = '" & charFCID & "' AND NUMNOMCONC = " & numNomConc & " AND NUMQCLEVEL = " & int12 & " AND CHARTYPE = 'Accuracy'"
        Dim rowsBias() As DataRow = tblQCTables.Select(strF)
        'get average
        Dim CountA As Short
        Dim numM As Decimal = 0
        Dim intM As Short = 0
        For CountA = 0 To rowsBias.Length - 1
            var1 = NZ(rowsBias(CountA).Item("numValue"), "")
            If Len(var1) = 0 Or IsNumeric(var1) = False Then
            Else
                intM = intM + 1
                numM = numM + CDec(var1)
            End If
        Next
        If intM = 0 Then
            GetBiasFromDiffCol = -1
        Else
            GetBiasFromDiffCol = SigFigOrDec(numM / intM, LSigFig, False) 'this does either sigfigs or decimal
        End If


    End Function

    Function GetGroup(ANALYTEDESCRIPTION_C As String)

        '20181128 LEE:

        Dim strF As String = "ANALYTEDESCRIPTION_C = '" & CleanText(ANALYTEDESCRIPTION_C) & "'"
        Dim row() As DataRow = tblAnalyteGroups.Select(strF)

        If row.Length = 0 Then
            GetGroup = 1
        Else
            GetGroup = row(0).Item("INTGROUP")
        End If


    End Function

    Function GetARSRuns(ByVal tblRID As DataTable, ByVal intAnalyteID As Int64, ByVal strAnalC As String, ByVal boolSA As Boolean) As String

        '20171123 LEE: Redo this
        'don't need to loop through each Run ID - takes a long time for large studies

        '20181128 LEE:
        'Hmmm. If there is more than one calibr level, must also evaluate LLOQ and ULOQ

        '20190305 LEE:
        'Regression Contstants tables can now have samples assigned
        'If boolSA (samples assigned), then ignore B column logic of dv1

        GetARSRuns = ""

        Dim dv1 As System.Data.DataView
        dv1 = frmH.dgvAnalyticalRunSummary.DataSource
        Dim tblARS As System.Data.DataTable = dv1.ToTable
        Dim strFFF As String
        Dim intFFF As Short = 0
        Dim strFARS As String
        Dim rowsARS() As DataRow
        Dim intARS As Int16
        Dim Count1 As Int32
        Dim Count2 As Int16
        Dim var1, var2, var3
        Dim str1 As String

        'make sure the runid's are unique
        Dim dv2 As DataView = New DataView(tblRID)
        Dim tblURID As DataTable = dv2.ToTable("a", True, "RUNID")

        'debug
        Dim intAAAA As Int64 = tblARS.Rows.Count
        intAAAA = intAAAA

        If Len(strAnalC) = 0 Then

            '20181015 LEE:
            'Solved!! ANALYTEID in tblARS is string!
            strFARS = "ANALYTEID = '" & intAnalyteID & "' AND BOOLINCLUDEREGR = " & False

        Else

            'in this case, Matrix is being passed, not strAnalC
            Try
                strFARS = "Matrix = '" & strAnalC & "' AND BOOLINCLUDEREGR = " & False
            Catch ex As Exception
                var1 = var1
            End Try

        End If

        Erase rowsARS
        Try
            Try
                rowsARS = tblARS.Select(strFARS)
            Catch ex As Exception
                '20181015 LEE:
                'See above for better explanation
                '20171120 LEE: 
                'If tblARS gets too large, SELECT has a problem with BOOLINCLUDEREGR
                'don't know why
                'must do a two-step filter

                'debug
                Dim strA As String
                Dim strB As String
                Dim strC As String
                Dim strD As String

                'strA = "[Watson Run ID] = '" & var1 & "'"
                strB = "ANALYTEID = " & intAnalyteID
                strC = "BOOLINCLUDEREGR = " & False
                strD = "BOOLINCLUDEREGR = 'False'"

                Try
                    Dim rowsAA() As DataRow
                    Dim tblAAA As DataTable
                    Dim rowsBBB() As DataRow
                    Dim tblBBB As DataTable

                    Try
                        rowsBBB = tblARS.Select(strB)
                        tblBBB = rowsBBB.CopyToDataTable
                        Try
                            rowsARS = tblBBB.Select(strC)
                        Catch ex2 As Exception
                            var1 = var1
                            Try
                                rowsARS = tblBBB.Select(strD)
                            Catch ex3 As Exception
                                var1 = var1
                            End Try
                        End Try
                    Catch ex5 As Exception
                        var1 = var1
                    End Try

                Catch ex1 As Exception
                    var1 = var1 'debug
                End Try

            End Try

            intARS = rowsARS.Length

            ''debug
            'For Count1 = 0 To tblARS.Columns.Count - 1
            '    var1 = tblARS.Columns(Count1).ColumnName
            '    var1 = var1
            'Next

            For Count1 = 0 To intARS - 1

                var1 = rowsARS(Count1).Item("Watson Run ID")
                intFFF = intFFF + 1
                If intFFF = 1 Then
                    strFFF = "RUNID <> " & var1
                Else
                    strFFF = strFFF & " AND RUNID <> " & var1
                End If

            Next



        Catch ex As Exception
            var1 = var1 'debug
        End Try


        If intFFF = 0 Then
        Else
            GetARSRuns = "(" & strFFF & ")"
        End If

end1:

    End Function


    Function GetNETProvider() As String

        GetNETProvider = constrIni

        Dim int1 As Short = InStr(1, constrIni, "DATA SOURCE", CompareMethod.Text)
        If int1 > 0 Then
            GetNETProvider = Mid(constrIni, int1, Len(constrIni))
        End If

    End Function

    Function StudyAllowed() As Boolean

        StudyAllowed = True

        Dim boolFL As Boolean = boolFormLoad

        If boolAccess Then
            GoTo end1
        End If

        Dim strM As String = ""
        Dim strF As String
        Dim wUserID As Int64
        Dim dgv As DataGridView = frmH.dgvwStudy
        Dim intRow As Int32
        Dim strStudy As String
        Dim intwStudyID As Int64
        Dim var1
        Dim strWA As String 'Watson Account

        If dgv.RowCount = 0 Then
            GoTo end1
        End If

        intRow = dgv.CurrentRow.Index

        Dim intORow As Int32 = frmH.txtcbxMDBSelIndex.Text

        Try

            strF = "ID_TBLUSERACCOUNTS = " & id_tblUserAccounts
            Dim rowsU() As DataRow = tblUserAccounts.Select(strF)
            wUserID = NZ(rowsU(0).Item("ID_TBLWATSONACCOUNT"), -1)

            If wUserID = 0 Or wUserID = -1 Then
                StudyAllowed = True
                'this user not assigned a Watson ID
                GoTo end1
            End If

            Dim strF1 As String
            intwStudyID = dgv("STUDYID", intRow).Value
            strF1 = "STUDYID = " & intwStudyID & " AND USERID = " & wUserID
            Dim rowsW() As DataRow = tblWatsonStudyRoles.Select(strF1)

            'get Watson User Account name
            Dim strF2 As String
            strF2 = "USERID = " & wUserID
            Dim rowsWA() As DataRow = tblWatsonUsers.Select(strF2)
            strWA = rowsWA(0).Item("LOGINNAME")

            If rowsW.Length = 0 Then
                StudyAllowed = False
                strStudy = dgv("STUDYNAME", intRow).Value
                strM = "Current StudyDoc user linked to Watson user account:" & ChrW(10) & ChrW(10) & "    '" & strWA & "'" & ChrW(10) & ChrW(10) & "does not have access to the the chosen study:"
                strM = strM & ChrW(10) & ChrW(10)
                strM = strM & "     " & strStudy
            Else
                StudyAllowed = True
            End If

            If StudyAllowed Then
            Else
                MsgBox(strM, vbInformation, "Invalid choice...")

                boolFormLoad = True
                If intORow = -10 Then

                    If dgv.RowCount > 0 Then
                        'select first row
                        dgv.CurrentCell = dgv.Item("STUDYNAME", 0)
                        'now clear selection
                        dgv.ClearSelection()
                    End If
                Else
                    Try
                        dgv.CurrentCell = dgv.Item("STUDYNAME", intORow)
                        dgv.CurrentRow.Selected = True
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try

                End If

            End If

        Catch ex As Exception

        End Try

end1:

        boolFormLoad = boolFL

    End Function

    Function GetISFromAnalyte(intI As Short) As String

        GetISFromAnalyte = "NA"

        Dim Count6 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String


        Dim tbl1 As System.Data.DataTable = tblAnalytesHome
        'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX

        Try
            'GetISFromAnalyte = NZ(tbl1.Rows(intI).Item("IntStd"), "NA")
            GetISFromAnalyte = NZ(tbl1.Rows(intI).Item("CHARUSERIS"), "NA")
        Catch ex As Exception

        End Try




    End Function

    Function FindAnalyteFromIntStdDtbl(ByVal intRowsAnal As Short, ByVal dtbl As System.Data.DataTable, ByVal strAnalyte As String) As Short

        FindAnalyteFromIntStdDtbl = -1

        Dim Count6 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        For Count6 = 0 To intRowsAnal - 1
            str1 = dtbl.Rows(Count6).Item("UseIntStd")
            str2 = dtbl.Rows(Count6).Item("IntStd")
            str3 = dtbl.Rows(Count6).Item("AnalyteDescription")
            If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                If StrComp(str2, strAnalyte, CompareMethod.Text) = 0 Then
                    If UseAnalyte(str3) Then
                        FindAnalyteFromIntStdDtbl = Count6
                        Exit For
                    End If
                End If
            End If
        Next


    End Function

    Function FindAnalyteFromIntStd(ByVal intRowsAnal As Short, ByVal rowS() As System.Data.DataRow, ByVal strAnalyte As String) As Short

        FindAnalyteFromIntStd = -1

        Dim Count6 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        For Count6 = 0 To intRowsAnal - 1
            str1 = rowS(Count6).Item("UseIntStd")
            str2 = rowS(Count6).Item("IntStd")
            str3 = rowS(Count6).Item("AnalyteDescription")
            If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                If StrComp(str2, strAnalyte, CompareMethod.Text) = 0 Then
                    If UseAnalyte(str3) Then
                        FindAnalyteFromIntStd = Count6
                        Exit For
                    End If
                End If
            End If
        Next


    End Function

    Function DoIntStdDtbl(ByVal intRowsAnal As Short, ByVal dtbl As System.Data.DataTable, ByVal strAnalyte As String) As Boolean

        DoIntStdDtbl = False

        Dim Count6 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1

        Try
            For Count6 = 0 To intRowsAnal - 1
                str1 = dtbl.Rows(Count6).Item("UseIntStd")
                str2 = dtbl.Rows(Count6).Item("IntStd")
                str3 = dtbl.Rows(Count6).Item("AnalyteDescription")
                If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                    If StrComp(str2, strAnalyte, CompareMethod.Text) = 0 Then
                        If UseAnalyte(str3) Then
                            DoIntStdDtbl = True
                            Exit For
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try



    End Function

    Function DoIntStd(ByVal intRowsAnal As Short, ByVal rowS() As System.Data.DataRow, ByVal strAnalyte As String) As Boolean

        DoIntStd = False

        Dim Count6 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1

        Try
            For Count6 = 0 To intRowsAnal - 1
                str1 = rowS(Count6).Item("UseIntStd")
                str2 = rowS(Count6).Item("IntStd")
                str3 = rowS(Count6).Item("AnalyteDescription")
                If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                    If StrComp(str2, strAnalyte, CompareMethod.Text) = 0 Then
                        If UseAnalyte(str3) Then
                            DoIntStd = True
                            Exit For
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try



    End Function

    Function ReturnDiffLabel() As String

        ReturnDiffLabel = "%Diff"

        If boolSTATSBIAS Then
            ReturnDiffLabel = "%Bias"
        ElseIf boolSTATSDIFF Then
            ReturnDiffLabel = "%Diff"
        ElseIf BOOLSTATSRE Then
            ReturnDiffLabel = "%RE"
        ElseIf boolTHEORETICAL Then
            ReturnDiffLabel = "%Theor"
        End If

    End Function

    Function CleanText(ByVal strA As String)

        '20190206 LEE
        'need a function to update examples of text that aren't possible in tbl.select statement

        Dim var1
        CleanText = strA
        Try
            CleanText = Replace(strA, "'", "''", 1, -1, CompareMethod.Text)
        Catch ex As Exception
            var1 = var1
        End Try

    End Function

    Function boolNeedsFC(ByVal strFC As String, strDesc As String, ByVal id1 As Int64, ByRef dtbl As System.Data.DataTable) As Boolean

        boolNeedsFC = False
        Dim strF As String
        '20190206 LEE:
        Dim strDescR As String = CleanText(strDesc) ' Replace(strDesc, "'", "''", 1, -1, CompareMethod.Text)

        Try


            'newRow("CHARDESCRIPTION") = strC

            'If id1 > 0 Then
            '    strF = "CHARFIELDCODE = '" & strFC & "' AND CHARDESCRIPTION = '" & strDescR & "' AND ID_TBLREPORTTABLE = " & id1
            'Else
            '    strF = "CHARFIELDCODE = '" & strFC & "' AND CHARDESCRIPTION = '" & strDescR & "'"
            'End If

            If id1 > 0 Then
                strF = "CHARFIELDCODE = '" & strFC & "' AND ID_TBLREPORTTABLE = " & id1
            Else
                strF = "CHARFIELDCODE = '" & strFC & "'"
            End If

            Dim rows() As System.Data.DataRow = dtbl.Select(strF)

            If rows.Length = 0 Then
                boolNeedsFC = True
            Else
                'check description to see if it needs to be modified
                strDescR = rows(0).Item("CHARDESCRIPTION")
                If StrComp(strDescR, strDesc, CompareMethod.Text) = 0 Then
                Else
                    rows(0).BeginEdit()
                    rows(0).Item("CHARDESCRIPTION") = strDesc
                    rows(0).EndEdit()
                End If

            End If

        Catch ex As Exception

            Dim var1
            var1 = ex.Message
        End Try


    End Function

    Function GetStudyDocHeader(boolShort As Boolean) As String

        'keep for later
        'GetStudyDocHeader = "LABIntegrity StudyDoc" & ChrW(8482) & " Study Design and Report Writing Manager"

        If boolShort Then
            GetStudyDocHeader = "Report Writing Manager"
        Else
            GetStudyDocHeader = "LABIntegrity StudyDoc" & ChrW(8482) & " Report Writing Manager"
        End If

    End Function

    Function ReturnStdQC(ByVal strX As String) As String

        ReturnStdQC = strX

        If BOOLUSESTDCOLLABELS Then
        Else
            GoTo end1
        End If

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim strDash As String
        Dim varI

        Dim Count1 As Short

        For Count1 = 1 To 5
            Select Case Count1
                Case 1
                    str1 = "LLOQ"
                Case 2
                    str1 = "Low"
                Case 3
                    str1 = "Mid"
                Case 4
                    str1 = "High"
                Case 5
                    str1 = "Dil"
            End Select

            int1 = InStr(1, strX, str1, CompareMethod.Text)

            If int1 > 0 Then

                If StrComp(str1, "Mid", CompareMethod.Text) = 0 Then
                    'check for Mid-1, Mid-2, etc
                    'see if next character is '-'
                    int3 = int1 + Len(str1)
                    strDash = Mid(strX, int3, 1)
                    If StrComp(strDash, "-", CompareMethod.Text) = 0 Then
                        varI = Mid(strX, int3 + 1, 1)
                        If IsNumeric(varI) Then
                            ReturnStdQC = "QC " & str1 & "-" & varI
                        Else
                            ReturnStdQC = "QC " & str1
                        End If
                    Else
                        ReturnStdQC = "QC " & str1
                    End If
                    'int2 = InStr(int1 + Len(str1), strX, "-", CompareMethod.Text)

                Else
                    ReturnStdQC = "QC " & str1
                End If


                Exit For
            End If

        Next

end1:

    End Function

    Function ShortSampleName(strX As String) As String

        ShortSampleName = strX

        Dim intA As Short
        Dim intB As Short
        intA = InStr(1, strX, " ", CompareMethod.Text)
        If intA = 0 Then
            GoTo end1
        End If
        intB = InStr(intA + 1, strX, " ", CompareMethod.Text)
        If intB = 0 Then
            GoTo end1
        End If
        ShortSampleName = Mid(strX, intB + 1, Len(strX))

end1:

    End Function

    Function StripNOT(strVal As String) As String

        StripNOT = strVal
        If Len(strVal) = 0 Then
            GoTo end1
        End If

        Dim intP1 As Short

        intP1 = InStr(1, strVal, "(", CompareMethod.Text)
        If intP1 > 0 Then
            If intP1 = 1 Then
                StripNOT = ""
            Else
                StripNOT = Mid(strVal, 1, intP1 - 1)
            End If
        End If

end1:

    End Function

    Function ReturnNOT(strVal As String, strColum As String) As String

        ReturnNOT = ""

        'this function will determine if autoassigned value contains NOT value
        'will return an array of OR values

        Dim intUB As Short = 0
        Dim v(2, 100) 'this is the positions in the string 
        '1=Start Pos, 2=Length of string
        Dim vS(100) 'This is actual strings

        If Len(strVal) = 0 Then
            GoTo end1
        End If

        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short

        Dim Count1 As Short

        Dim intP1 As Short
        Dim intP2 As Short

        Dim strNew As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strX As String

        Dim numOR As Short

        intP1 = InStr(1, strVal, "(", CompareMethod.Text)
        If intP1 = 0 Then
            GoTo end1
        End If

        intP2 = InStr(intP1 + 1, strVal, ")", CompareMethod.Text)
        If intP2 - intP1 <= 1 Then
            GoTo end1
        End If

        'get new string
        Dim strFind As String = " OR "
        strNew = Mid(strVal, intP1 + 1, intP2 - intP1 - 1)

        'find number of OR's
        int1 = InStr(1, strNew, strFind, CompareMethod.Text)
        If int1 = 0 Then 'no OR's
            intUB = 1
            ReturnNOT = "(" & strColum & " NOT LIKE '*" & strNew & "*')"
        Else

            intUB = intUB + 1
            v(1, intUB) = 1
            v(2, intUB) = int1 - 1

            int2 = int1 + Len(strFind)
            Do Until int1 = 0

                int1 = InStr(int2, strNew, strFind, CompareMethod.Text)
                If int1 = 0 Then
                    Exit Do
                End If

                intUB = intUB + 1
                v(1, intUB) = int2
                int3 = int1 - int2
                v(2, intUB) = int3

                int2 = int1 + Len(strFind)

            Loop

            'get last value
            intUB = intUB + 1
            v(1, intUB) = int2
            int3 = Len(strNew) - int2 + 1
            v(2, intUB) = int3

            For Count1 = 1 To intUB
                int1 = v(1, Count1)
                int2 = v(2, Count1)
                str1 = Trim(Mid(strNew, int1, int2))
                If Count1 = 1 Then
                    ReturnNOT = "(" & strColum & " NOT LIKE '*" & str1 & "*'"
                ElseIf Count1 = intUB Then
                    ReturnNOT = ReturnNOT & " AND " & strColum & " NOT LIKE '*" & str1 & "*')"
                Else
                    ReturnNOT = ReturnNOT & " AND " & strColum & " NOT LIKE '*" & str1 & "*'"
                End If
            Next

        End If


end1:

    End Function

    Function HasLogic(ByVal strL As String) As Boolean

        HasLogic = False

        Dim Count1 As Short
        Dim str1 As String

        For Count1 = 1 To 3

            Select Case Count1
                Case 1
                    str1 = " AND "
                Case 2
                    str1 = " OR "
                Case 3
                    str1 = " NOT("
            End Select

            'search for string
            If InStr(strL, str1, CompareMethod.Text) > 0 Then
                HasLogic = True
                Exit For
            End If

        Next

    End Function


    Function PivotASP(oldTable As DataTable, Optional pivotColumnOrdinal As Integer = 0) As DataTable

        Dim newTable As New DataTable
        Dim dr As DataRow
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim row As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim strType As String
        Dim strF As String
        Dim int1 As Int32
        Dim int2 As Int32
        Dim int3 As Int32
        Dim var1, var2

        int1 = oldTable.Rows.Count 'debug

        'add two columns
        For Count1 = 1 To 9

            str2 = ""

            Select Case Count1
                Case Is = 1
                    str1 = "CHARLABEL"
                    strType = "System.String"
                    str2 = "Text Fragment Type"
                Case Is = 2
                    str1 = "CHARVALUE"
                    strType = "System.String"
                    str2 = "Text Fragment"

                Case Is = 3
                    str1 = "CHARNOT"
                    strType = "System.String"
                    str2 = "NOT"

                Case Is = 4
                    str1 = "CHAREXAMPLE"
                    strType = "System.String"
                    str2 = "Example"

                Case Is = 5
                    str1 = "CHARCOLUMNNAME"
                    strType = "System.String"
                    str2 = "COLUMNNAME"

                Case Is = 6
                    str1 = "ID_TBLSTUDIES"
                    strType = "System.Int64"
                Case Is = 7
                    str1 = "ID_TBLCONFIGREPORTTABLES"
                    strType = "System.Int64"
                Case Is = 8
                    str1 = "ID_TBLREPORTTABLE"
                    strType = "System.Int64"

                Case Is = 9
                    str1 = "INTORDER"
                    strType = "System.Int16"

            End Select

            Dim col1 As New DataColumn
            col1.ColumnName = str1
            col1.DataType = System.Type.GetType(strType)
            col1.Caption = str2
            'col1.DefaultValue = 0
            col1.AllowDBNull = True
            newTable.Columns.Add(col1)

        Next

        ' loop through columns
        'add rows for each row in study 44
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim rows() As DataRow = tblAutoAssignSamples.Select(strF)

        ''debug
        'For Count1 = 0 To tblTableProperties.Columns.Count - 1
        '    'console.writeline(tblTableProperties.Columns(Count1).ColumnName)
        'Next

        Dim intOrder As Short = 1
        Dim intCt As Short = 0

        For Count1 = 0 To rows.Length - 1

            int1 = rows(Count1).Item("ID_TBLSTUDIES")
            int2 = rows(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            int3 = rows(Count1).Item("ID_TBLREPORTTABLE")

            For Count2 = 0 To oldTable.Columns.Count - 1

                var1 = oldTable.Columns(Count2).ColumnName
                ' each column becomes a new row
                dr = newTable.NewRow()

                dr.BeginEdit()
                dr.Item("CHARLABEL") = GetCaptionA(var1.ToString) ' oldTable.Columns(Count2).Caption

                Select Case var1
                    Case "ID_TBLSTUDIES"
                        dr.Item("CHARVALUE") = int1
                    Case "ID_TBLCONFIGREPORTTABLES"
                        dr.Item("CHARVALUE") = int2
                    Case "ID_TBLREPORTTABLE"
                        dr.Item("CHARVALUE") = int3

                    Case Is = "BOOLUSESTDCOLLABELS"
                        intOrder = 1
                    Case Is = "CHARRECPES"
                        intOrder = 2
                    Case Is = "CHARRECRS"
                        intOrder = 3
                    Case Is = "CHARRECQC"
                        intOrder = 4
                    Case Is = "CHARLOT1"
                        intOrder = 5
                    Case Is = "CHARLOT2"
                        intOrder = 6
                    Case Is = "CHARLOT3"
                        intOrder = 7
                    Case Is = "CHARLOT4"
                        intOrder = 8
                    Case Is = "CHARLOT5"
                        intOrder = 9
                    Case Is = "CHARLOT6"
                        intOrder = 10
                    Case Is = "CHARLOT7"
                        intOrder = 11
                    Case Is = "CHARLOT8"
                        intOrder = 12
                    Case Is = "CHARLOT9"
                        intOrder = 13
                    Case Is = "CHARLOT10"
                        intOrder = 14

                        '20181216 LEE:
                        'added WOIS lots
                    Case Is = "CHARLOTWOIS1"
                        intOrder = 15
                    Case Is = "CHARLOTWOIS2"
                        intOrder = 16
                    Case Is = "CHARLOTWOIS3"
                        intOrder = 17
                    Case Is = "CHARLOTWOIS4"
                        intOrder = 18
                    Case Is = "CHARLOTWOIS5"
                        intOrder = 19
                    Case Is = "CHARLOTWOIS6"
                        intOrder = 20
                    Case Is = "CHARLOTWOIS7"
                        intOrder = 21
                    Case Is = "CHARLOTWOIS8"
                        intOrder = 22
                    Case Is = "CHARLOTWOIS9"
                        intOrder = 23
                    Case Is = "CHARLOTWOIS10"
                        intOrder = 24




                    Case Is = "CHARDILN"
                        intOrder = 25
                    Case Is = "CHARDILNFACTOR"
                        intOrder = 26
                    Case Is = "CHARCALSTD"
                        intOrder = 27
                    Case Is = "CHARNONCOREQC"
                        intOrder = 28
                    Case Is = "CHARCOREQC"
                        intOrder = 29
                    Case Is = "CHAROLD"
                        intOrder = 30
                    Case Is = "CHARNEW"
                        intOrder = 31
                    Case Is = "CHARNEW2"
                        intOrder = 32
                    Case Is = "CHARNEW3"
                        intOrder = 33



                    Case Is = "CHARLLOQ"
                        intOrder = 34
                    Case Is = "CHARULOQ"
                        intOrder = 35
                    Case Is = "CHARBLANK"
                        intOrder = 36
                    Case Is = "CHARSTOCKSOLNCONC"
                        intOrder = 37
                    Case Is = "CHARRUNIDENTIFIER1"
                        intOrder = 38
                    Case Is = "CHARRUNIDENTIFIER2"
                        intOrder = 39
                    Case Is = "CHARRUNIDENTIFIER3"
                        intOrder = 40
                    Case Is = "CHARRUNIDENTIFIER4"
                        intOrder = 41

                    Case Is = "CHARSAMPLETYPE"
                        intOrder = 42
                    Case Is = "CHARRUNDESCR1"
                        intOrder = 43
                    Case Is = "CHARRUNDESCR2"
                        intOrder = 44
                    Case Is = "BOOLACCEPTEDONLY"
                        intOrder = 45
                    Case Else
                End Select


                dr.Item("CHARCOLUMNNAME") = var1

                dr.Item("ID_TBLSTUDIES") = int1
                dr.Item("ID_TBLCONFIGREPORTTABLES") = int2
                dr.Item("ID_TBLREPORTTABLE") = int3
                dr.Item("INTORDER") = intOrder

                dr.EndEdit()

                'add the DataRow to the new table
                newTable.Rows.Add(dr)
            Next

        Next

        int1 = newTable.Rows.Count 'debug

        Return newTable

    End Function

    Function ReturnGrey(strC As String, boolNot As Boolean) As Boolean

        ReturnGrey = False

        Dim str1 As String = ""
        Dim str2 As String = "Sample Name text: "
        Dim str3 As String
        Dim boolAdd2 As Boolean = True

        Select Case strC

            Case Is = "BOOLUSESTDCOLLABELS"
                str1 = "Use Standard QC Label Conventions"
                boolAdd2 = False
                ReturnGrey = True
            Case Is = "CHARRECPES"
                str1 = "Post-Extraction Spike Solution"

            Case Is = "CHARRECRS"
                str1 = "Recovery Solution"

            Case Is = "CHARRECQC"
                str1 = "Extracted QC Standard"

            Case Is = "CHARLOT1"
                str1 = "Lot 1"

            Case Is = "CHARLOT2"
                str1 = "Lot 2"

            Case Is = "CHARLOT3"
                str1 = "Lot 3"

            Case Is = "CHARLOT4"
                str1 = "Lot 4"

            Case Is = "CHARLOT5"
                str1 = "Lot 5"

            Case Is = "CHARLOT6"
                str1 = "Lot 6"

            Case Is = "CHARLOT7"
                str1 = "Lot 7"

            Case Is = "CHARLOT8"
                str1 = "Lot 8"

            Case Is = "CHARLOT9"
                str1 = "Lot 9"

            Case Is = "CHARLOT10"
                str1 = "Lot 10"

            Case Is = "CHARDILN"
                str1 = "Dilution QC Samples"

            Case Is = "CHARDILNFACTOR"
                str1 = "Dilution QC Samples Dilution Factor (optional)"
                boolAdd2 = False
                ReturnGrey = True

            Case Is = "CHARCALSTD"
                str1 = "Calibration Standards"

            Case Is = "CHARNONCOREQC"
                str1 = "Non Intra/Inter Run QC"

            Case Is = "CHAROLD"
                str1 = "Old or First Measurement"

            Case Is = "CHARNEW"
                str1 = "New or Second Measurement"

            Case Is = "CHARLLOQ"
                str1 = "LLOQ"

            Case Is = "CHARULOQ"
                str1 = "ULOQ"

            Case Is = "CHARBLANK"
                str1 = "Blank"

            Case Is = "CHARSTOCKSOLNCONC"
                str1 = "Stock Solution Concentration (Optional)"
                boolAdd2 = False
                ReturnGrey = True

            Case Is = "CHARRUNIDENTIFIER1"
                str1 = "Run Identifier 1 (Required)"
                boolAdd2 = False
                ReturnGrey = True

            Case Is = "CHARRUNIDENTIFIER2"
                str1 = "Run Identifier 2 (Required)"
                boolAdd2 = False
                ReturnGrey = True


            Case Is = "CHARRUNIDENTIFIER3"
                str1 = "Run Identifier 3 (Optional)"
                boolAdd2 = False
                ReturnGrey = True

            Case Is = "CHARRUNIDENTIFIER4"
                str1 = "Run Identifier 4 (Optional)"
                boolAdd2 = False
                ReturnGrey = True


            Case Is = "CHARCOREQC"
                str1 = "Intra/Inter Run QC"

            Case Is = "CHARSAMPLETYPE"
                str1 = "Sample Type"
                boolAdd2 = False
                ReturnGrey = True

            Case Is = "CHARRUNDESCR1"
                str1 = "Text contained in Analytical Run Description"
                boolAdd2 = False

            Case Is = "CHARRUNDESCR2"
                str1 = "Text contained in Analytical Run Description 2"
                boolAdd2 = False

            Case Is = "BOOLACCEPTEDONLY"
                str1 = "Query ONLY analytical runs with Accepted regression"
                boolAdd2 = False
                ReturnGrey = True

        End Select

    End Function

    Function GetCaptionA(strC As String) As String

        GetCaptionA = strC

        Dim str1 As String = ""
        Dim str2 As String = "Sample Name text: "
        Dim str3 As String
        Dim boolAdd2 As Boolean = True

        Select Case strC

            Case Is = "BOOLUSESTDCOLLABELS"
                str1 = "Use Standard QC Label Conventions"
                boolAdd2 = False

            Case Is = "CHARRECPES"
                str1 = "Post-Extraction Spike Solution"

            Case Is = "CHARRECRS"
                str1 = "Recovery Solution"

            Case Is = "CHARRECQC"
                str1 = "Extracted QC Standard"

            Case Is = "CHARLOT1"
                str1 = "Lot 1"

            Case Is = "CHARLOT2"
                str1 = "Lot 2"

            Case Is = "CHARLOT3"
                str1 = "Lot 3"

            Case Is = "CHARLOT4"
                str1 = "Lot 4"

            Case Is = "CHARLOT5"
                str1 = "Lot 5"

            Case Is = "CHARLOT6"
                str1 = "Lot 6"

            Case Is = "CHARLOT7"
                str1 = "Lot 7"

            Case Is = "CHARLOT8"
                str1 = "Lot 8"

            Case Is = "CHARLOT9"
                str1 = "Lot 9"

            Case Is = "CHARLOT10"
                str1 = "Lot 10"


            Case Is = "CHARLOTWOIS1"
                str1 = "Lot 1"

            Case Is = "CHARLOTWOIS2"
                str1 = "Lot 2"

            Case Is = "CHARLOTWOIS3"
                str1 = "Lot 3"

            Case Is = "CHARLOTWOIS4"
                str1 = "Lot 4"

            Case Is = "CHARLOTWOIS5"
                str1 = "Lot 5"

            Case Is = "CHARLOTWOIS6"
                str1 = "Lot 6"

            Case Is = "CHARLOTWOIS7"
                str1 = "Lot 7"

            Case Is = "CHARLOTWOIS8"
                str1 = "Lot 8"

            Case Is = "CHARLOTWOIS9"
                str1 = "Lot 9"

            Case Is = "CHARLOTWOIS10"
                str1 = "Lot 10"



            Case Is = "CHARDILN"
                str1 = "Dilution QC Samples"

            Case Is = "CHARDILNFACTOR"
                str1 = "Dilution QC Samples Dilution Factor (optional)"
                boolAdd2 = False

            Case Is = "CHARCALSTD"
                str1 = "Calibration Standards"

            Case Is = "CHARNONCOREQC"
                str1 = "Non Intra/Inter Run QC"

            Case Is = "CHAROLD"
                str1 = "Old or First Measurement"

            Case Is = "CHARNEW"
                str1 = "New or Second Measurement"
            Case Is = "CHARNEW2"
                str1 = "New or Third Measurement"
            Case Is = "CHARNEW3"
                str1 = "New or Fourth Measurement"

            Case Is = "CHARLLOQ"
                str1 = "LLOQ"

            Case Is = "CHARULOQ"
                str1 = "ULOQ"

            Case Is = "CHARBLANK"
                str1 = "Blank"

            Case Is = "CHARSTOCKSOLNCONC"
                str1 = "Stock Solution Concentration (Optional)"
                boolAdd2 = False

            Case Is = "CHARRUNIDENTIFIER1"
                str1 = "Run Identifier 1 (Required)"
                boolAdd2 = False

            Case Is = "CHARRUNIDENTIFIER2"
                str1 = "Run Identifier 2 (Required)"
                boolAdd2 = False

            Case Is = "CHARRUNIDENTIFIER3"
                str1 = "Run Identifier 3 (Optional)"
                boolAdd2 = False
            Case Is = "CHARRUNIDENTIFIER4"
                str1 = "Run Identifier 4 (Optional)"
                boolAdd2 = False


            Case Is = "CHARCOREQC"
                str1 = "Intra/Inter Run QC"

            Case Is = "CHARSAMPLETYPE"
                str1 = "Sample Type"
                boolAdd2 = False

            Case Is = "CHARRUNDESCR1"
                str1 = "Text contained in Analytical Run Description"
                boolAdd2 = False

            Case Is = "CHARRUNDESCR2"
                str1 = "Text contained in Analytical Run Description 2"
                boolAdd2 = False

            Case Is = "BOOLACCEPTEDONLY"
                str1 = "Query ONLY analytical runs with Accepted regression"
                boolAdd2 = False

        End Select

        If Len(str1) = 0 Then
        Else
            If boolAdd2 Then
                str3 = str2 & str1
            Else
                str3 = str1
            End If
            GetCaptionA = str3
        End If


    End Function

    Function HasSpecialCharacters(strC As String) As Short

        'This function is used in ReportTableConfig - AutoAssign columns cell validating

        HasSpecialCharacters = 0

        '        These are valid characters:

        '[space]
        'a -z
        'A -Z
        '0-9
        '-
        '_
        '.
        '/
        '\

        'http://stackoverflow.com/questions/3701018/remove-special-characters-from-a-string
        'Dim cleanString As String = Regex.Replace(yourString, "[^A-Za-z0-9\-/]", "")
        Dim cs As String ' = Regex.Replace(strC, "[^A-Za-z0-9\-_. ]", "")
        cs = Regex.Replace(strC, "[^A-Za-z0-9\-_./\\# ]", "")
        'Dim cs As String = Regex.Replace(strC, "[^A-Za-z0-9\-_. ]", "")

        Dim int1 As Short
        'int1 = Math.Abs(Len(cs) - Len(strC))
        'HasSpecialCharacters = int1

        '20160929 LEE: change this logic to allow anything except apostrophe
        If InStr(1, strC, "'", CompareMethod.Text) Then
            HasSpecialCharacters = 1
        End If

    End Function

    Function ReturnRecoveryLabel(intTableID As Int64, intTCHL As Int64)

        ReturnRecoveryLabel = ""

        Dim str1 As String
        Dim strF As String
        Dim intColsH As Short
        'get intcols from tblreportdata
        Dim tblH As System.Data.DataTable
        Dim rowsH() As DataRow

        tblH = tblReportTableHeaderConfig
        strF = "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND BOOLINCLUDE <> 0 AND ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGHEADERLOOKUP = " & intTCHL
        rowsH = tblH.Select(strF)
        intColsH = rowsH.Length
        If intColsH = 0 Then
            ReturnRecoveryLabel = ""
        Else
            ReturnRecoveryLabel = NZ(rowsH(0).Item("CHARUSERLABEL"), "")
        End If

    End Function

    Sub SelectRows(tbl As Word.Table, row1 As Int16, row2 As Int16)

        Dim rowRange As Word.Range
        With tbl
            rowRange = .Rows(row1).Range
            rowRange.End = .Rows(row2).Range.End
        End With
        rowRange.Select()

    End Sub

    Function HasISRPDifference(strAnalyteID As String, idTR As Int64) As Boolean

        HasISRPDifference = False

        Dim rowsCrit() As DataRow
        Dim rowsID() As DataRow
        Dim ID1 As Decimal
        Dim strFC As String
        Dim intIncl As Short

        'first find id
        strFC = "ID_TBLCONFIGREPORTTABLES = 30 AND CHARCOLUMNLABEL = '%Difference'"
        rowsID = tblConfigHeaderLookup.Select(strFC)
        ID1 = rowsID(0).Item("ID_TBLCONFIGHEADERLOOKUP")

        'now get table specific value
        strFC = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 30 AND ID_TBLCONFIGHEADERLOOKUP = " & ID1
        rowsCrit = tblReportTableHeaderConfig.Select(strFC)
        intIncl = NZ(rowsCrit(0).Item("BOOLINCLUDE"), 0)
        If intIncl = 0 Then
            HasISRPDifference = False
        Else
            HasISRPDifference = True
        End If

    End Function

    Function HasISRPassFail(strAnalyteID As String, idTR As Int64) As Boolean

        HasISRPassFail = False

        Dim rowsCrit() As DataRow
        Dim rowsID() As DataRow
        Dim ID1 As Decimal
        Dim strFC As String
        Dim intIncl As Short

        'first find id
        strFC = "ID_TBLCONFIGREPORTTABLES = 30 AND CHARCOLUMNLABEL = 'PASS/FAIL'"
        rowsID = tblConfigHeaderLookup.Select(strFC)
        ID1 = rowsID(0).Item("ID_TBLCONFIGHEADERLOOKUP")

        'now get table specific value
        strFC = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 30 AND ID_TBLCONFIGHEADERLOOKUP = " & ID1
        rowsCrit = tblReportTableHeaderConfig.Select(strFC)
        intIncl = NZ(rowsCrit(0).Item("BOOLINCLUDE"), 0)
        If intIncl = 0 Then
            HasISRPassFail = False
        Else
            HasISRPassFail = True
        End If

    End Function

    Function ReturnISRCrit1(strAnalyteID As String, idTR As Int64) As Decimal

        Dim rowsCrit() As DataRow
        Dim numCrit As Decimal
        Dim strFC As String
        strFC = "ANALYTEID = " & strAnalyteID & " AND ID_TBLREPORTTABLE = " & idTR
        rowsCrit = tblReportTableAnalytes.Select(strFC)
        If rowsCrit.Length = 0 Then
            numCrit = 20
        Else
            numCrit = NZ(rowsCrit(0).Item("NUMINCSAMPLECRIT01"), 20)
        End If
        ReturnISRCrit1 = numCrit

    End Function

    Sub SetvWatsonDB()

        'DBTABLEVER: 7.4
        'DBQUERYVER: 5.4
        'JETVER: 
        'VBVER: 
        'KEYFIELD: 
        'JETBITS: 
        'VBBITS: 
        'ENVIRONMENT: WATP
        'WATSONVER: 7.4.1

        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim Count1 As Short
        Dim var1

        Try

            'Dim boolJustTable As Boolean
            strWatsonVersion = tblWatsonDBVersion.Rows(tblWatsonDBVersion.Rows.Count - 1).Item("WATSONVER")
            strWatsonDBVersion = tblWatsonDBVersion.Rows(tblWatsonDBVersion.Rows.Count - 1).Item("DBTABLEVER")
            'convert to number
            Dim v() = GetChar(strWatsonDBVersion, ".")

            Dim iA As Short
            '1=pos
            int1 = UBound(v)
            int3 = 0
            For Count1 = 1 To int1
                int3 = int3 + 1
                If Count1 = 1 Then
                    int2 = 1
                End If
                iA = v(Count1)
                var1 = Mid(strWatsonDBVersion, int2, iA - int2)
                vWatsonDB(Count1) = CInt(var1)
                int2 = iA + 1
            Next
            If int3 = 1 Then
                var1 = Mid(strWatsonDBVersion, iA + 1, Len(strWatsonDBVersion))
                vWatsonDB(2) = CInt(var1)
                vWatsonDB(3) = 0
            Else
                var1 = Mid(strWatsonDBVersion, iA + 1, Len(strWatsonDBVersion))
                vWatsonDB(3) = CInt(var1)
            End If

            strSDWatsonDBVErsion = Format(vWatsonDB(1), "00") & Format(vWatsonDB(2), "00") & Format(vWatsonDB(3), "00")

            If strSDWatsonDBVErsion >= "070400" Then
                boolCanDoISR = True
            Else
                boolCanDoISR = False
            End If

        Catch ex As Exception

        End Try


    End Sub

    Function GetChar(x As String, ch As String)

        '1=pos,2=total#
        Dim v(100)
        Dim Count1 As Short

        Dim cnt As Integer = 0
        Dim intLX As Short = Len(x)
        Dim c As String

        For Count1 = 1 To intLX
            c = Mid(x, Count1, 1)
            If StrComp(c, ch, CompareMethod.Text) = 0 Then
                cnt = cnt + 1
                v(cnt) = Count1

            End If
        Next

        ReDim Preserve v(cnt)

        GetChar = v

        Return GetChar

    End Function

    Function ReturnDate(dt As Date) As String

        'cstr(dt) = '5/27/2016 4:44:10 PM'

        If boolGuWuSQLServer Then
            ReturnDate = "'" & CStr(dt) & "'"
        ElseIf boolGuWuAccess Then
            ReturnDate = "#" & CStr(dt) & "#"
        ElseIf boolGuWuOracle Then
            'TO_DATE('2015/05/15 8:30:25', 'YYYY/MM/DD HH:MI:SS')
            'ReturnDate = "TO_DATE(" & CStr(dt) & ", 'MM/DD/YYYY HH:MI:SS AM')"
            'or TO_DATE('1998-DEC-25 17:30','YYYY-MON-DD HH24:MI:SS'
            'This may not be correct
            ReturnDate = "TO_DATE('" & CStr(dt) & "', 'YYYY/MM/DD HH:MI:SS AM')"
        End If

    End Function

    Sub TestOutsideAccCrit()

        'for testing

        Dim numLo As Decimal
        Dim numHi As Decimal

        Dim numConc As Decimal = 144
        Dim numNomConc As Decimal = 125

        Dim numCrit1 As Decimal = 15
        Dim numCrit2 As Decimal = 15

        Dim num1 As Decimal
        Dim num2 As Decimal
        Dim num3 As Decimal
        Dim num4 As Decimal

        numLo = numNomConc - (numNomConc * numCrit1 / 100)
        num1 = numLo

        numHi = numNomConc + (numNomConc * numCrit1 / 100)
        num2 = numHi

        num3 = SigFigOrDec(num2, 3, False)
        num3 = num3

        Dim boolO As Boolean
        boolO = OutsideAccCrit(numConc, numNomConc, numCrit1, numCrit2, False)
        boolO = boolO


    End Sub

    Function OutsideAccCrit(ByVal varConc As Object, ByVal numNomConc As Decimal, ByVal crit1 As Decimal, ByVal crit2 As Decimal, ByVal intUseGuWuAccCrit As Short) As Boolean

        '20160820 LEE: Note that this function essentially negates the need for Sub SetHighAndLowCriteria
        'This function also negates the need for the rounding choice 'Criteria Precision Convention'

        'intUseGuWuAccCrit: 0 = false, -1 = true
        OutsideAccCrit = False

        '20180430 LEE
        If numNomConc = 0 Then
            GoTo end1
        End If

        Dim numLo As Decimal
        Dim numHi As Decimal

        Dim num1 As Decimal
        Dim num2 As Decimal
        Dim numConc As Decimal

        'must check to ensure numConc is not null
        '20180430 LEE:
        'must account for Endogenous Cmpds, NomConc = 0
        If IsDBNull(varConc) Or varConc <= 0 Then
            GoTo end1
        End If

        numConc = CDec(varConc)

        'LDEC: configured decimal places to round

        num1 = Math.Abs(((numConc / numNomConc) - 1) * 100)
        num2 = RoundToDecimal(num1, LDec)

        If intUseGuWuAccCrit <> 0 And gAllowGuWuAccCrit And LAllowGuWuAccCrit Then 'animal health allows asynchronous lo/hi
            If numConc <= numNomConc Then
                If num2 > crit1 Then
                    OutsideAccCrit = True
                End If
            Else
                If num2 > crit2 Then
                    OutsideAccCrit = True
                End If
            End If

        Else

            If num2 > crit1 Then
                OutsideAccCrit = True
            End If

        End If

end1:

    End Function

    Function GetVisibleCol(dgv As DataGridView) As Short

        GetVisibleCol = 0

        Dim Count1 As Short
        'find 1st visible column
        For Count1 = 0 To dgv.ColumnCount - 1
            If dgv.Columns(Count1).Visible Then
                dgv.CurrentCell = dgv.Item(Count1, 0)
                Exit For
            End If
        Next

    End Function

    Function GetMaxID(ByVal strF As String, ByVal intIncr As Int16, ByVal boolIncr As Boolean) As Int64

        '20190218 LEE: Frontage reports crashing when 5 users attempt to load 5 new studies and apply template simultaneously.
        '"Cannot insert duplicate key in object 'dbo.TBLREPORTTABLE'. The duplicate key value is 5126
        'This looks like a maxID thing. 
        'We're going to have to use GetMaxID
        'Modify function to pass incrementer

        '20190306 LEE: If boolIncr = true, then increment value

        Dim tblmax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim maxID As Int64
        Dim strFMax As String
        Dim var1

        'first refresh maxid tables
        If boolGuWuOracle Then
            ta_tblMaxID.Fill(tblMaxID)
        ElseIf boolGuWuAccess Then
            ta_tblMaxIDAcc.Fill(tblMaxID)
        ElseIf boolGuWuSQLServer Then
            ta_tblMaxIDSQLServer.Fill(tblMaxID)
        End If

        strFMax = "charTable = '" & strF & "'"
        tblmax = tblMaxID
        Try
            rowsMax = tblmax.Select(strFMax)
        Catch ex As Exception
            var1 = var1
        End Try

        maxID = rowsMax(0).Item("NUMMAXID")
        maxID = maxID + 1
        GetMaxID = maxID
        If boolIncr Then
            rowsMax(0).BeginEdit()
            rowsMax(0).Item("NUMMAXID") = maxID + intIncr
            rowsMax(0).EndEdit()
            Try
                If boolGuWuOracle Then
                    Try
                        ta_tblMaxID.Update(tblMaxID)
                    Catch ex As DBConcurrencyException
                        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                    End Try
                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblMaxIDAcc.Update(tblMaxID)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                    End Try
                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblMaxIDSQLServer.Update(tblMaxID)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                    End Try
                End If

            Catch ex As Exception

                Dim table As Data.DataTable
                Dim row As Data.DataRow
                Dim int1 As Int16

                If tblMaxID.HasErrors Then

                    int1 = 0
                    For Each row In tblMaxID.Rows
                        int1 = int1 + 1
                        If row.HasErrors Then
                            var1 = row.RowError
                            var1 = var1
                            var1 = row.RowState
                            var1 = var1
                            ' Process error here. 
                        End If
                    Next
                End If

                var1 = ex.Message
                var1 = var1
            End Try
        End If

        var1 = var1


    End Function

    Function PutMaxID(ByVal strF As String, ByVal maxID As Int64) As Boolean

        Dim tblmax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim var1

        strFMax = "charTable = '" & strF & "'"
        tblmax = tblMaxID
        rowsMax = tblmax.Select(strFMax)
        rowsMax(0).BeginEdit()
        rowsMax(0).Item("nummaxid") = maxID
        rowsMax(0).EndEdit()

        Try
            If boolGuWuOracle Then
                Try
                    ta_tblMaxID.Update(tblMaxID)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblMaxIDAcc.Update(tblMaxID)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblMaxIDSQLServer.Update(tblMaxID)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                End Try
            End If


        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


    End Function

    Function GetRunID(intAssayID As Int64) As Int16

        Dim strF As String
        strF = "ASSAYID = " & intAssayID
        Dim rows() As DataRow = tblCalStdGroupAssayIDsAll.Select(strF)
        GetRunID = rows(0).Item("RUNID")

    End Function

    Function GetANALYTEFLAGPERCENTAnova(numNomConc As Single, intRunID As Int16, intAnalyteID As Int64, ByRef tblLevelCrit As DataTable) As Single

        'if nomConc < 1, then the query return 0 records
        'must do something different

        Dim strF As String
        Dim var1, var2
        Dim Count4 As Short
        Dim boolHit As Boolean

        strF = "ANALYTEID = " & intAnalyteID & " AND RUNID = " & intRunID

        Dim intRC As Short
        Dim rows10() As DataRow = tblQCRunIDs.Select(strF, "LEVELNUMBER ASC")
        intRC = rows10.Length
        var1 = 15
        If intRC = 0 Then
            var1 = 15
        Else
            Dim num1 As Single
            boolHit = False
            For Count4 = 0 To intRC - 1
                num1 = NZ(rows10(Count4).Item("CONCENTRATION"), -1)
                'varNom is double, must convert to single
                If num1 = numNomConc Then
                    var2 = NZ(rows10(Count4).Item("FLAGPERCENT"), 15)
                    var1 = NZ(rows10(Count4).Item("ANALYTEFLAGPERCENT"), var2)
                    boolHit = True
                    Exit For
                End If
            Next
            If boolHit Then
                strF = "NomConc = " & numNomConc & " AND Crit = " & var1
                Dim rows() = tblLevelCrit.Select(strF)
                If rows.Length = 0 Then
                    Dim nr As DataRow = tblLevelCrit.NewRow
                    nr.BeginEdit()
                    nr.Item("NomConc") = numNomConc
                    nr.Item("Crit") = var1
                    nr.EndEdit()
                    tblLevelCrit.Rows.Add(nr)
                End If
            Else
                strF = "NomConc = " & numNomConc
                Dim rows() As DataRow = tblLevelCrit.Select(strF)
                If rows.Length = 0 Then

                Else
                    var1 = rows(0).Item("Crit")
                End If
            End If
        End If

        GetANALYTEFLAGPERCENTAnova = CSng(var1)

    End Function

    Function GetANALYTEFLAGPERCENT(numNomConc As Single, intRunID As Int16, intAnalyteID As Int64) As Single

        'if nomConc < 1, then the query return 0 records
        'must do something different

        Dim strF As String
        Dim var1, var2
        Dim Count4 As Short

        strF = "ANALYTEID = " & intAnalyteID & " AND RUNID = " & intRunID

        Dim intRC As Short
        Dim rows10() As DataRow = tblQCRunIDs.Select(strF, "LEVELNUMBER ASC")
        intRC = rows10.Length
        var1 = 15
        If intRC = 0 Then
            var1 = 15
        Else
            Dim num1 As Single
            For Count4 = 0 To intRC - 1
                num1 = NZ(rows10(Count4).Item("CONCENTRATION"), -1)
                'varNom is double, must convert to single
                If num1 = numNomConc Then
                    var2 = NZ(rows10(Count4).Item("FLAGPERCENT"), 15)
                    var1 = NZ(rows10(Count4).Item("ANALYTEFLAGPERCENT"), var2)
                    Exit For
                End If
            Next
        End If

        GetANALYTEFLAGPERCENT = CSng(var1)

    End Function

    Function GetConcUnits(intRunID As Int16) As String

        GetConcUnits = ""

        Dim strF As String
        strF = "RUNID = " & intRunID
        Dim rows() As DataRow = tblCalStdGroupAssayIDsAll.Select(strF)
        If rows.Length = 0 Then
        Else
            GetConcUnits = rows(0).Item("CONCENTRATIONUNITS")
        End If

    End Function

    Function ReturnSort(boolGroupsTable As Boolean) As String

        If boolGroupsTable Then
            If StrComp(gSortAnalytes, "Matrix", CompareMethod.Text) = 0 Then
                ReturnSort = "MATRIX ASC, ANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION_C ASC, INTGROUP ASC"
            Else
                ReturnSort = "ANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION_C ASC, MATRIX ASC, INTGROUP ASC"
            End If
        Else
            If StrComp(gSortAnalytes, "Matrix", CompareMethod.Text) = 0 Then
                ReturnSort = "MATRIX ASC, ORIGINALANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION ASC, INTGROUP ASC"
            Else
                ReturnSort = "ORIGINALANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION ASC, MATRIX ASC, INTGROUP ASC"
            End If
        End If


    End Function

    Function UpdateGroupsAssignedSamples() As Boolean

        UpdateGroupsAssignedSamples = True

        '20160226 LEE: The implementation of Groups poses a problem with StudyDoc database table tblAssignedSamples. 
        'tblAssignedSamples records ANALYTEINDEX, but unfortunately, doesn't record ANALYTEID - which is vital in the Groups model
        'Previously column INTGROUP had been added to tblAssignedSamples and given a default value of 0
        'With StudyDoc 3.0.27.5, column INTANALYTEID has been added to tblAssignedSamples
        'This subroutine UpdateGroupsAssignedSamples will fire everytime a Watson study has been open.
        'It will inspect tblAssignedSamples (filtered for the appropriate study) for any records where INTGROUP = 0 and update them with the appropriate INTANALYTEID and INTGROUP

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strF As String
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim var1, var2, var3
        Dim intRunID As Int16
        Dim intAnalIndex As Int32
        Dim intAnalID As Int32
        Dim intGroup As Short
        Dim boolIS As Boolean
        Dim intIS As Short '0=no, -1=yes

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND (INTGROUP = 0 OR INTGROUP IS NULL)"

        Dim rowsAS() As DataRow = tblAssignedSamples.Select(strF)

        Dim intRows As Int32
        intRows = rowsAS.Length

        If intRows = 0 Then
            UpdateGroupsAssignedSamples = False
            GoTo end1
        End If

        Dim rowsAAR() As DataRow
        Dim rowsCS() As DataRow
        Try
            For Count1 = 0 To intRows - 1
                intRunID = NZ(rowsAS(Count1).Item("RUNID"), 0)
                intAnalIndex = NZ(rowsAS(Count1).Item("ANALYTEINDEX"), 0)
                intIS = rowsAS(Count1).Item("BOOLINTSTD") '0=no, -1=yes
                If intIS = 0 Then
                    boolIS = False
                Else
                    boolIS = True
                End If

                'get analyteid from tblallanalruns
                strF = "ANALYTEINDEX = " & intAnalIndex & " AND RUNID = " & intRunID
                Erase rowsAAR
                rowsAAR = tblAllAnalRuns.Select(strF)
                If rowsAAR.Length = 0 Then
                    var1 = var1 'debug
                Else
                    'retrieve analyteid
                    intAnalID = rowsAAR(0).Item("ANALYTEID")

                    'tblAnalysisResultsHome
                    'now get intGroup from tblCalStdGroupAssayIDsAll
                    strF = "ANALYTEINDEX = " & intAnalIndex & " AND RUNID = " & intRunID & " AND ANALYTEID = " & intAnalID
                    Erase rowsCS
                    rowsCS = tblCalStdGroupAssayIDsAll.Select(strF)
                    If rowsCS.Length = 0 Then
                        var1 = var1
                    Else
                        intGroup = rowsCS(0).Item("INTGROUP")
                        'now record analyteid and intgroup
                        rowsAS(Count1).BeginEdit()
                        rowsAS(Count1).Item("INTANALYTEID") = intAnalID
                        If boolIS Then
                            rowsAS(Count1).Item("INTGROUP") = -1
                        Else
                            rowsAS(Count1).Item("INTGROUP") = intGroup
                        End If

                        rowsAS(Count1).EndEdit()
                    End If
                End If
            Next

            'update to the database
            If boolGuWuOracle Then
                Try
                    ta_tblAssignedSamples.Update(tblAssignedSamples)
                Catch ex As DBConcurrencyException
                    ds2005.TBLASSIGNEDSAMPLES.Merge(ds2005.TBLASSIGNEDSAMPLES, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblAssignedSamplesAcc.Update(tblAssignedSamples)
                Catch ex As DBConcurrencyException
                    ds2005Acc.TBLASSIGNEDSAMPLES.Merge(ds2005Acc.TBLASSIGNEDSAMPLES, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblAssignedSamplesSQLServer.Update(tblAssignedSamples)
                Catch ex As DBConcurrencyException
                    ds2005Acc.TBLASSIGNEDSAMPLES.Merge(ds2005Acc.TBLASSIGNEDSAMPLES, True)
                End Try
            End If

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


end1:


    End Function

    Function GetASSAYIDFilter(intGroup As Short, boolAll As Boolean, boolFromUseGroups As Boolean) As String

        Dim Count2 As Int16
        Dim str1 As String
        Dim str2 As String
        Dim strF As String

        GetASSAYIDFilter = ""

        str1 = "INTGROUP = " & intGroup
        'don't need strFAssayID here

        'get strFAssayID to possibly use in the future
        Dim rowsAllRuns() As DataRow
        '20160818 LEE:
        'New logic to bring in Not Regressed runs
        If boolAll Then
            rowsAllRuns = tblCalStdGroupAssayIDsAll.Select("INTGROUP = " & intGroup)
        Else
            If boolFromUseGroups Then
                rowsAllRuns = tblCalStdGroupAssayIDsAcc.Select("INTGROUP = " & intGroup)
            Else
                strF = " AND (RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3)"
                rowsAllRuns = tblCalStdGroupAssayIDsAll.Select(str1 & strF)
            End If

        End If
        If rowsAllRuns.Length = 0 Then
            GoTo end1
        End If

        GetASSAYIDFilter = "(ASSAYID = "
        For Count2 = 0 To rowsAllRuns.Length - 1
            str1 = rowsAllRuns(Count2).Item("ASSAYID").ToString
            If Count2 = 0 Then
                GetASSAYIDFilter = GetASSAYIDFilter & str1
            Else
                GetASSAYIDFilter = GetASSAYIDFilter & " OR ASSAYID = " & str1
            End If
        Next
        If rowsAllRuns.Length = 0 Then
            GetASSAYIDFilter = "(ASSAYID = -1)"
        Else
            GetASSAYIDFilter = GetASSAYIDFilter & ")" 'add parenthesis because will be used later as part of future filters
        End If


end1:

    End Function

    Function GetASSAYIDFilterIDCT(intGroup As Short, boolAll As Boolean, boolFromUseGroups As Boolean, idCT As Int64) As String

        Dim Count2 As Int16
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim var1, var2

        '20170105 LEE: Must also take into account Anal Run Review selection and if QC or Calibr Stds
        Dim boolANR As Boolean = False
        Select Case idCT
            Case 2, 3, 4, 11 'Regr, Calibr, QC,QCAnova
                boolANR = True
        End Select

        Dim dvANR As DataView = frmH.dgvAnalyticalRunSummary.DataSource
        Dim tblANR As DataTable = dvANR.ToTable
        Dim strFANR As String

        GetASSAYIDFilterIDCT = ""

        str1 = "INTGROUP = " & intGroup
        'don't need strFAssayID here

        'get strFAssayID to possibly use in the future
        Dim rowsAllRuns() As DataRow
        '20160818 LEE:
        'New logic to bring in Not Regressed runs
        If boolAll Then
            rowsAllRuns = tblCalStdGroupAssayIDsAll.Select("INTGROUP = " & intGroup)
        Else
            If boolFromUseGroups Then
                rowsAllRuns = tblCalStdGroupAssayIDsAcc.Select("INTGROUP = " & intGroup)
            Else
                strF = " AND (RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3)"
                rowsAllRuns = tblCalStdGroupAssayIDsAll.Select(str1 & strF)
            End If

        End If
        If rowsAllRuns.Length = 0 Then
            GoTo end1
        End If

        GetASSAYIDFilterIDCT = "(ASSAYID = "
        Dim intF As Short = 0
        Dim boolF As Boolean
        Dim intAID As Short
        For Count2 = 0 To rowsAllRuns.Length - 1
            str1 = rowsAllRuns(Count2).Item("ASSAYID").ToString
            intAID = rowsAllRuns(Count2).Item("RUNID")
            boolF = True
            If boolANR Then 'check AnalRunReview
                'tblANR does not have ASSAYID
                strFANR = "[Watson Run ID] = '" & intAID & "' AND boolIncludeRegr = TRUE"
                Dim rowsANR() As DataRow
                Try
                    rowsANR = tblANR.Select(strFANR)
                Catch ex As Exception
                    var1 = ex.Message
                End Try
                If rowsANR.Length = 0 Then
                    boolF = False
                End If
            End If

            If boolF Then
                intF = intF + 1
                If intF = 1 Then
                    GetASSAYIDFilterIDCT = GetASSAYIDFilterIDCT & str1
                Else
                    GetASSAYIDFilterIDCT = GetASSAYIDFilterIDCT & " OR ASSAYID = " & str1
                End If
            End If
        Next
        If intF = 0 Then
            GetASSAYIDFilterIDCT = "(ASSAYID = -1)"
        Else
            GetASSAYIDFilterIDCT = GetASSAYIDFilterIDCT & ")" 'add parenthesis because will be used later as part of future filters
        End If


end1:

    End Function

    Function GetNumMatrices() As Short

        'find number of species
        Dim dvSP As System.Data.DataView = New DataView(tblSpeciesMatrix)
        Dim tblSP As System.Data.DataTable = dvSP.ToTable("aSP", True, "SPECIES")
        intNumSpecies = tblSP.Rows.Count

        'find number of matrixes
        tblSP = dvSP.ToTable("aSM", True, "SAMPLETYPEID")
        GetNumMatrices = tblSP.Rows.Count

    End Function

    Function GetUserAnalyteNameNoGroup(strA As String) As String

        GetUserAnalyteNameNoGroup = strA

        Dim strF As String
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND (AnalyteDescription = '" & CleanText(strA) & "' OR ORIGINALANALYTEDESCRIPTION = '" & CleanText(strA) & "')"
        Dim rows() As DataRow = TBLSTUDYDOCANALYTES.Select(strF)
        If rows.Length = 0 Then
        Else

            Dim str1 As String
            Dim str2 As String
            str1 = rows(0).Item("IsIntStd")
            If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                'now find Analyte
                Dim strF1 As String
                strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND IntStd = '" & CleanText(strA) & "'"
                Dim rows1() As DataRow = TBLSTUDYDOCANALYTES.Select(strF1)
                Dim Count1 As Int16
                For Count1 = 0 To rows.Length - 1
                    str2 = NZ(rows1(Count1).Item("CHARUSERIS"), "")
                    If Len(str2) = 0 Then
                    Else
                        GetUserAnalyteNameNoGroup = str2
                        Exit For
                    End If
                Next
                'GetUserAnalyteNameNoGroup = rows(0).Item("CHARUSERIS")
            Else
                GetUserAnalyteNameNoGroup = rows(0).Item("CHARUSERANALYTE")
            End If

        End If
     

end1:

    End Function

    Function GetUserAnalyteName(strA As String, boolX As Boolean, intGroup As Int16) As String

        GetUserAnalyteName = strA

        Dim strF As String
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim rows() As DataRow
        If intGroup < 1 Or boolX Then
            'is IntStd
            'look for CHARUSERANALYTE in original analyte
            strF = strF & " AND IsIntStd = 'No' AND IntStd = '" & CleanText(strA) & "' AND IntStd IS NOT NULL"
            rows = TBLSTUDYDOCANALYTES.Select(strF)
            If rows.Length = 0 Then
            Else
                GetUserAnalyteName = NZ(rows(0).Item("CHARUSERIS"), strA)
            End If
        Else
            strF = strF & " AND INTGROUP = " & intGroup
            rows = TBLSTUDYDOCANALYTES.Select(strF)
            If rows.Length = 0 Then
            Else
                GetUserAnalyteName = NZ(rows(0).Item("CHARUSERanalyte"), strA)
            End If
        End If

end1:

    End Function

    Function UpdateAnalyteMatrix(strT As String, strAA As String, strM As String, boolDoC As Boolean, intGroup As Short, boolX As Boolean) As String

        'Updates [ANALYTE] and [MATRIX] in Table Header using individual matrix values.
        'Necessary for multi-matrix studies.

        '20181108 LEE:
        'strA must be updated to CHARUSERANALYTE
        'boolX = True = is IntStd
        Dim strA As String
        strA = GetUserAnalyteName(strAA, boolX, intGroup)

        Dim varAReplace, varMReplace, varMReplaceLC

        UpdateAnalyteMatrix = strT

        'Adjust Analyte Replacement Text (as in modStyle2:AddTableNumber)
        If IsDBNull(strA) Then
            varAReplace = "[NA]"
        ElseIf Len(strA) = 0 Then
            varAReplace = "[NA]"
        Else
            'varReplace = LowerCase(var8) 'UnCapit(Trim(var8), True)
            varAReplace = Capit(UnCapit(strA, False))
            varAReplace = Replace(varAReplace, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
        End If

        UpdateAnalyteMatrix = Replace(UpdateAnalyteMatrix, "[ANALYTE]", strA, 1, -1, CompareMethod.Text)

        '20181015 LEE:
        'put nbh back in
        UpdateAnalyteMatrix = Replace(UpdateAnalyteMatrix, "-", NBHReal, 1, -1, CompareMethod.Text)

        'Adjust Matrix Replacement Text (as in modStyle2:AddTableNumber)
        If IsDBNull(strM) Then
            varMReplace = "[NA]"
        ElseIf Len(strM) = 0 Then
            varMReplace = "[NA]"
        Else
            varMReplaceLC = LCase(Trim(strM))
            varMReplace = Capit(LCase(Trim(strM)))
            varMReplace = Replace(varMReplace, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
        End If

        'UpdateAnalyteMatrix = Replace(UpdateAnalyteMatrix, "[MATRIX]", strM, 1, -1, CompareMethod.Text)
        UpdateAnalyteMatrix = Replace(UpdateAnalyteMatrix, "[LC_MATRIX]", varMReplaceLC, 1, -1, CompareMethod.Text)
        UpdateAnalyteMatrix = Replace(UpdateAnalyteMatrix, "[UC_MATRIX]", varMReplace, 1, -1, CompareMethod.Text)
        UpdateAnalyteMatrix = Replace(UpdateAnalyteMatrix, "[MATRIX]", varMReplace, 1, -1, CompareMethod.Text)

        If boolDoC And BOOLCALIBRTABLETITLE Then

            Dim strCR As String 'calibration range
            Dim var1, var2, var3, var4
            Dim strLLOQ As String
            Dim strULOQ As String
            Dim strUnits As String
            Dim intAnalyteID As Int64
            Dim strF As String
            strF = "INTGROUP = " & intGroup

            'first get analyteid from tblAnalyteGroupsAcc: accepted groups
            Dim rowsAG() As DataRow = tblAnalyteGroupsAcc.Select(strF)
            If rowsAG.Length = 0 Then
                GoTo end1
            End If
            'get analyteid
            intAnalyteID = rowsAG(0).Item("ANALYTEID")

            'now see if there is more than one analyteid in tblAnalyteGroupsAcc
            strF = "ANALYTEID = " & intAnalyteID & " AND MATRIX ='" & strM & "'"
            Dim rowsIDs() As DataRow = tblAnalyteGroupsAcc.Select(strF)
            If rowsIDs.Length < 2 Then
                GoTo end1
            End If

            'now filter tblCalStdGroupsAcc for intgroup and matrix
            strF = "INTGROUP = " & intGroup
            Dim rows() As DataRow = tblCalStdGroupsAcc.Select(strF)
            If rows.Length = 0 Then
            Else
                var1 = rows(0).Item("LLOQ")
                If Len(NZ(var1, "")) = 0 Then
                    GoTo end1
                End If
                'convert to sigfigs
                var1 = SigFigOrDecString(var1, LSigFig, False)
                strCR = "(" & var1 & " to "

                var2 = rows(0).Item("ULOQ")
                If Len(NZ(var2, "")) = 0 Then
                    GoTo end1
                End If
                var2 = SigFigOrDecString(var2, LSigFig, False)
                strCR = strCR & var2

                var3 = rows(0).Item("CONCENTRATIONUNITS")
                If Len(NZ(var3, "")) = 0 Then
                    GoTo end1
                End If
                strCR = strCR & " " & var3 & ")"

                'replace spaces with NBS
                strCR = Replace(strCR, " ", ChrW(160), 1, -1, CompareMethod.Text)

                'add space in front of strCR
                strCR = " " & strCR

                'add strCR to UpdateAnalyteMatrix
                UpdateAnalyteMatrix = UpdateAnalyteMatrix & strCR

            End If

            'tblCalStdGroupsAcc
        End If

end1:

    End Function

    Function ReturnTableTitle(intIndex As Short) As String

        Dim strA

        Dim strAnal As String = tblAnalyteGroups.Rows(intIndex).Item("ANALYTEDESCRIPTION")
        Dim strAnalC As String = tblAnalyteGroups.Rows(intIndex).Item("ANALYTEDESCRIPTION_C")
        Dim strMatrix As String = tblAnalyteGroups.Rows(intIndex).Item("MATRIX")

        If intNumMatrix = 1 Then
            strA = strAnalC 'arrAnalytes(14, Count1)
        Else
            strA = strAnal & " in " & strMatrix
        End If

    End Function

    Function ReturnLeftIndent(wd As Microsoft.Office.Interop.Word.Application, boolAppendix As Boolean, boolAttachment As Boolean, boolTable As Boolean) As Single

        ReturnLeftIndent = 72

        Dim str1 As String
        Dim var1
        Dim numLI As Single 'Left Indent
        'set left indent depending on font size and font type
        'The current selection is 'caption' style

        'minimum of 72

        If boolAppendix Then
            numLI = 72 ' 1 inch
        ElseIf boolAttachment Then
            numLI = 81 ' 1 1/8 inch
        ElseIf boolTable Then
            numLI = 72 ' 54 ' 3/4 inch
        End If

        With wd
            str1 = .Selection.Font.Name
            var1 = .Selection.Font.Size

            If InStr(1, str1, "Arial", CompareMethod.Text) > 0 Then

                If boolAppendix Then
                    If var1 >= 12 Then
                        numLI = 81 '1 1/8 inch
                    Else
                        numLI = 72 ' 1 inch
                    End If
                ElseIf boolAttachment Then
                    If var1 >= 12 Then
                        numLI = 90 '1 1/4 inch
                    Else
                        numLI = 81 ' 1 1/8 inch
                    End If
                ElseIf boolTable Then
                    If var1 >= 12 Then
                        numLI = 72 ' 63 '7/8 inch
                    Else
                        numLI = 72 ' 54 ' 3/4 inch
                    End If
                End If

            End If
        End With

        ReturnLeftIndent = numLI

    End Function

    Function boolAnalyte_C() As Boolean

        boolAnalyte_C = False

        Dim Count1 As Short
        Dim str1 As String
        Dim tbl As DataTable = tblAnalyteGroups

        For Count1 = 0 To tbl.Rows.Count - 1

            str1 = NZ(tbl.Rows(Count1).Item("ANALYTEDESCRIPTION_C"), "")
            If InStr(1, str1, "_C1", CompareMethod.Text) > 0 Then
                boolAnalyte_C = True
                Exit For
            End If

        Next


    End Function

    Function boolMultiMatrix() As Boolean

        boolMultiMatrix = False

        Try

            Dim Count1 As Short
            Dim str1 As String
            Dim tbl As DataTable = tblSpeciesMatrix

            Dim dv As DataView = New DataView(tblSpeciesMatrix, "", "", DataViewRowState.CurrentRows)
            Dim tbl1 As DataTable = dv.ToTable("a", True, "SAMPLETYPEID")

            If tbl1.Rows.Count = 1 Then
                boolMultiMatrix = False
            Else
                boolMultiMatrix = True
            End If
        Catch ex As Exception

        End Try

    End Function

    Function GetUserID() As String

        GetUserID = "Guest"

        Dim tblU As DataTable
        Dim strF As String
        Dim rowU() As DataRow
        Dim str1 As String

        tblU = tblUserAccounts

        'find user account
        strF = "id_tblUserAccounts = " & idU ' id_tblUserAccounts
        rowU = tblU.Select(strF)
        If rowU.Length = 0 Then
            GetUserID = "Guest" ' rowU(0).Item("charUserID")
        Else
            GetUserID = rowU(0).Item("charUserID")
        End If


    End Function

    Function IBS(strP As String) As String 'Insert Back Slash

        IBS = strP

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short

        int1 = Len(strP)

        str1 = Mid(strP, int1, 1)

        If StrComp(str1, "\", CompareMethod.Text) = 0 Then
        Else
            IBS = strP & "\"
        End If

    End Function

    Function ApplyReportTemplate(frm As Form, charModule As String) As Boolean

        ApplyReportTemplate = False

        Dim var1, var2, var3
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim MaxID As Long
        Dim strF As String
        Dim strO As String
        Dim ctl As Windows.Forms.Control ' Control
        Dim str1 As String
        Dim str2 As String
        Dim boolOpt As Boolean
        Dim boolChk As Boolean
        Dim boolCbx As Boolean
        Dim boolStr As Boolean
        Dim con As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim boolGo As Boolean
        Dim gbx As Windows.Forms.GroupBox
        Dim chk As Windows.Forms.CheckBox

        Dim int1 As Short
        Dim int2 As Short

        Dim dtbl As DataTable = TBLSECTIONTEMPLATES


        strF = "charModule = '" & charModule & "'"
        Dim rows() As DataRow = dtbl.Select(strF, strO)
        int1 = rows.Length
        If int1 = 0 Then
            int1 = 0
            GoTo end1
        End If

        Select Case charModule

            Case "Report Prelim"

                'has 1 groupboxe full of checkboxes
                For Count1 = 1 To 1
                    Select Case Count1
                        Case 1
                            gbx = frm.Controls("gb1")
                            strF = "charModule = '" & charModule & "' AND INTINDEX = 1"
                            rows = dtbl.Select(strF)
                    End Select

                    For Count2 = 0 To rows.Length - 1
                        str1 = NZ(rows(Count2).Item("CHARCONTROL"), "NA")
                        var1 = rows(Count2).Item("NUMVALUE")
                        For Each ctl In gbx.Controls
                            Try
                                'ctl = gbx.Controls(Count3) ' ctl
                                str2 = ctl.Name
                                If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                                    chk = ctl
                                    chk.Checked = var1
                                    Exit For
                                End If

                            Catch ex As Exception

                            End Try

                        Next
                    Next

                Next

                'now do form itself
                strF = "charModule = '" & charModule & "' AND INTINDEX = 0"
                rows = dtbl.Select(strF)
                For Count2 = 0 To rows.Length - 1
                    str1 = NZ(rows(Count2).Item("CHARCONTROL"), "NA")
                    var1 = rows(Count2).Item("NUMVALUE")

                    If StrComp(str1, "chkAdvSettings", CompareMethod.Text) = 0 Then
                        var1 = var1 'debug
                    End If

                    For Each ctl In frm.Controls
                        Try
                            chk = ctl
                            str2 = ctl.Name
                            If StrComp(str2, "chkAdvSettings", CompareMethod.Text) = 0 Then
                                var1 = var1 'debug
                            End If
                            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                                chk.Checked = var1
                                Exit For
                            End If

                        Catch ex As Exception

                        End Try

                    Next
                Next

next2:

            Case "Word Compare"

                'has 1 groupboxe full of checkboxes
                For Count1 = 1 To 1
                    Select Case Count1
                        Case 1
                            gbx = frm.Controls("gbSettings")
                            strF = "charModule = '" & charModule & "' AND INTINDEX = 1"
                            rows = dtbl.Select(strF)
                    End Select

                    For Count2 = 0 To rows.Length - 1
                        str1 = NZ(rows(Count2).Item("CHARCONTROL"), "NA")
                        var1 = rows(Count2).Item("NUMVALUE")
                        For Each ctl In gbx.Controls
                            Try
                                'ctl = gbx.Controls(Count3) ' ctl
                                str2 = ctl.Name
                                If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                                    chk = ctl
                                    chk.Checked = var1
                                    Exit For
                                End If

                            Catch ex As Exception

                            End Try

                        Next
                    Next

                Next

        End Select

end1:

        If int1 = 0 Then
            ApplyReportTemplate = False
        Else
            ApplyReportTemplate = True
        End If

    End Function

    Function SampleName(strA As String) As String

        SampleName = strA

        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Integer
        Dim int2 As Integer

        'find first set of SS
        int1 = InStr(1, strA, " ", vbTextCompare)
        If int1 < 1 Then
            GoTo end1
        End If
        str1 = Trim(Mid(strA, int1, Len(strA)))
        SampleName = str1

        'remove dilution factor from end
        For Count1 = Len(str1) To 1 Step -1
            str2 = Mid(str1, Count1, 1)
            If StrComp(str2, " ", vbTextCompare) = 0 Then
                SampleName = Trim(Mid(str1, 1, Count1))
                Exit For
            End If

        Next

end1:

    End Function

    Function RoundDown(ByVal var As Double) As Int16

        Dim var1, var2
        Dim int1 As Integer

        var1 = CInt(var)
        If var1 = CDec(var) Then
            RoundDown = var1
        Else
            'strip decimal
            int1 = InStr(1, CStr(var), ".", vbTextCompare)
            var2 = Left(var, int1 - 1)
            RoundDown = CInt(var2)
        End If


    End Function

    Function RoundUp(ByVal var As Double) As Int16

        Dim var1, var2, var3
        Dim int1 As Integer

        var1 = var + 1

        RoundUp = RoundDown(var1)

    End Function


    Function GetDefaultRFC() As String

        GetDefaultRFC = ""

        Dim tbl As System.Data.DataTable
        Dim tblC As System.Data.DataTable = tblConfigCompliance
        Dim rows() As DataRow
        Dim strF As String
        Dim str1 As String

        '****
        Dim intRFC As Short

        intRFC = tblConfigCompliance.Rows(0).Item("BOOLREASONFORCHANGE")

        If intRFC = 0 Then
            GetDefaultRFC = "[Reason For Change option disabled]"
        Else
            tbl = tblReasonForChange
            strF = "BOOLDEFAULT = -1"
            rows = tbl.Select(strF)

            If rows.Length = 0 Then
                GetDefaultRFC = "[Reason for Change default not configured]"
            Else
                GetDefaultRFC = rows(0).Item("CHARREASONFORCHANGE")
            End If
        End If

        '****

    End Function

    Function GetDefaultMOS() As String

        GetDefaultMOS = ""

        Dim tbl As System.Data.DataTable = tblMeaningOfSig
        Dim tblC As System.Data.DataTable = tblConfigCompliance
        Dim rows() As DataRow
        Dim strF As String

        '****
        Dim intMOS As Short

        intMOS = tblConfigCompliance.Rows(0).Item("BOOLMEANINGOFSIG")
        If intMOS = 0 Then
            GetDefaultMOS = "[Meaning of Signature option disabled]"
        Else
            strF = "BOOLDEFAULT = -1"
            rows = tbl.Select(strF)

            If rows.Length = 0 Then
                GetDefaultMOS = "[Meaning of Signature default not configured]"
            Else
                GetDefaultMOS = rows(0).Item("CHARMEANINGOFSIG")
            End If
        End If
        '****


    End Function



    Function AllowPrint() As Boolean

        If BOOLALLOWPDFREPORT = False And BOOLALLOWREPORTGENERATION = False Then
            MsgBox("This user does not have Report Generation privileges.", MsgBoxStyle.Information, "No no...")
            AllowPrint = False
        Else
            AllowPrint = True
        End If

    End Function

    Function RetrieveSG(ByVal ID_TBLCONFIGHEADERLOOKUP As Int64) As String

        Dim dtbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intRows1 As Short
        Dim strF1 As String
        Dim strS1 As String

        'ID_TBLCONFIGHEADERLOOKUP
        dtbl1 = tblConfigHeaderLookup

        If ID_TBLCONFIGHEADERLOOKUP = 0 Then
            RetrieveSG = "[None]"
        Else
            dtbl1 = tblConfigHeaderLookup
            strF1 = "ID_TBLCONFIGHEADERLOOKUP = " & ID_TBLCONFIGHEADERLOOKUP
            rows1 = dtbl1.Select(strF1)
            RetrieveSG = rows1(0).Item("CHARCOLUMNLABEL")
        End If

    End Function

    Function PutSG(ByVal CHARUSERLABEL As String) As Int64

        Dim dtbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intRows1 As Short
        Dim strF1 As String
        Dim strS1 As String

        dtbl1 = tblConfigHeaderLookup

        If StrComp(CHARUSERLABEL, "[None]", CompareMethod.Text) = 0 Then
            PutSG = 0
        Else
            dtbl1 = tblConfigHeaderLookup
            'strF1 = "CHARUSERLABEL = '" & CHARUSERLABEL & "'"
            'rows1 = dtbl1.Select(strF1)
            strF1 = "CHARCOLUMNLABEL = '" & CHARUSERLABEL & "' AND ID_TBLCONFIGREPORTTABLES = 5"
            rows1 = dtbl1.Select(strF1)
            PutSG = rows1(0).Item("ID_TBLCONFIGHEADERLOOKUP")
        End If

    End Function

    Function Coding(ByVal etxt As String, ByVal encrypt As Boolean) As String

        'To do a keyword encryption
        '(this is a simple one)
        'you just use the key as the algorithm

        'in this case its raw, real char ASCII + keyword char ASCII
        'plus you need to loop it, so 200+200=400 you -255 to correct it.
        'then its opposite to decode, just - instead od +
        'here's an example i wote for you.

        'The strings are converted to char arrays for performance sake, 
        'if you were to send a large string through it if it used strings, 
        'each iteration would be slower because of the (string = string + char) 
        'operation having to handle a larger value each time, unlike an array.
        '
        'k is the index of the key, it has to be iterated seperately to i because 
        'the key may be a smaller size than the text being encoded, so it needs to be able to loop
        'around back to 0

        Dim i As Integer
        Dim key As String
        key = "buendorf"
        Dim oldtext() As Char = etxt.ToCharArray
        Dim thekey() As Char = key.ToCharArray
        Dim newtxt(etxt.Length - 1) As Char
        Dim k As Integer
        For i = 0 To oldtext.Length - 1
            If encrypt = True Then
                newtxt(i) = Chr(FixIt(Asc(oldtext(i)) + Asc(thekey(k))))
            Else
                newtxt(i) = Chr(FixIt(Asc(oldtext(i)) - Asc(thekey(k))))
            End If
            k += 1
            If k = key.Length Then k = 0
        Next
        Return newtxt

    End Function


    Function GetWordVersion(id As Int64, boolTemplate As Boolean) As Int64

        'id = id_tblWordStatement
        Dim strF As String
        Dim strS As String
        Dim rows() As DataRow

        If boolTemplate Then
            strF = "ID_TBLWORDSTATEMENTS = " & id
            strS = "INTWORDVERSION DESC"

            rows = TBLWORDSTATEMENTSVERSIONS.Select(strF, strS, DataViewRowState.CurrentRows)

            If rows.Length = 0 Then
                GetWordVersion = 0
            Else
                GetWordVersion = rows(0).Item("INTWORDVERSION")
            End If
        Else
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARREPORTTYPE = 'Final Report'"
            strS = "INTFINALREPORTVERSION DESC"

            rows = tblFinalReport.Select(strF, strS, DataViewRowState.CurrentRows)

            If rows.Length = 0 Then
                GetWordVersion = 0
            Else
                GetWordVersion = rows(0).Item("INTFINALREPORTVERSION")
            End If
        End If



    End Function


    Function Createxml(id As Int64, intVersion As Int64) As String

        'save as temp then display in afr
        Dim strP As String
        strP = GetNewTempFile(True)
        strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)

        Dim intRow As Short
        Dim strPath As String
        'Dim dgv As DataGridView
        Dim strLbl As String


        Dim var1, var2
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim strW As String
        Dim fs As FileStream
        Dim strM As String
        'Dim id As Int64
        'Dim intVersion As Int64

        dtbl1 = tblWordStatements


        'first must open tblWordDocs

        Call OpenWordDocs(id, intVersion)

        dtbl2 = tblWorddocs

        strF = "ID_TBLWORDSTATEMENTS = " & id & " AND INTWORDVERSION = " & intVersion
        strS = "ID_TBLWORDDOCS ASC"
        rows2 = dtbl2.Select(strF, strS)
        intL = rows2.Length

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        strpathT = ""
        strpathT = GetNewTempFile(True)

        'why am i putting a .docx on it?
        'why not leave it as xml?
        'especially since the file is being built as xml
        'leave it as xml
        'strpathT = Replace(strpathT, ".xml", ".docx", 1, -1, CompareMethod.Text)

        strW = ""
        Dim strBuild As New StringBuilder("")

        For Count1 = 0 To intL - 1
            strBuild.Append(rows2(Count1).Item("CHARXML"))
        Next
        strW = strBuild.ToString()

        ' Add some information to the file.
        Dim info As Byte()
        If intL = 0 Then
            strM = "There is a problem with this data:" & ChrW(10)
            strM = strM & "tblWorddocs: " & strF & ChrW(10)
            strM = strM & "Please contact your StudyDoc system administrator."
            info = New UTF8Encoding(True).GetBytes(strM)
            strpathT = Replace(strpathT, ".XML", ".TXT", 1, -1, CompareMethod.Text)
        Else
            ' Add some information to the file.
            info = New UTF8Encoding(True).GetBytes(strW)
        End If

        fs = File.Create(strpathT)
        fs.Close()
        fs = File.OpenWrite(strpathT)

        fs.Write(info, 0, info.Length)
        fs.Close()

        Createxml = strpathT


    End Function

    Function CreatexmlHome(dgv As DataGridView) As String

        'save as temp then display in afr
        Dim strP As String
        strP = GetNewTempFile(True)
        strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)

        Dim intRow As Int32
        Dim strPath As String
        'Dim dgv As DataGridView
        Dim strLbl As String


        Dim var1, var2
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intL As Int64
        Dim strpathT As String
        Dim Count1 As Int16
        Dim strW As String
        Dim fs As FileStream
        Dim strM As String
        Dim id As Int64
        Dim intVersion As Int64

        dtbl1 = tblWordStatements

        If dgv.CurrentRow Is Nothing Then
            If dgv.RowCount = 0 Then
                GoTo end1
            Else
                dgv.Rows(0).Selected = True
            End If
        Else
            intRow = dgv.CurrentRow.Index
        End If

        'first must open tblWordDocs
        id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value
        intVersion = GetWordVersion(id, True)

        Call OpenWordDocs(id, intVersion)

        dtbl2 = tblWorddocs

        strF = "ID_TBLWORDSTATEMENTS = " & id & " AND INTWORDVERSION = " & intVersion
        strS = "ID_TBLWORDDOCS ASC"
        rows2 = dtbl2.Select(strF, strS)
        intL = rows2.Length

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        strpathT = ""
        strpathT = GetNewTempFile(True)

        'why am i putting a .docx on it?
        'why not leave it as xml?
        'especially since the file is being built as xml
        'leave it as xml
        'strpathT = Replace(strpathT, ".xml", ".docx", 1, -1, CompareMethod.Text)

        Dim strBuild = New StringBuilder()
        For Count1 = 0 To intL - 1
            strBuild.Append(rows2(Count1).Item("CHARXML"))
        Next

        ' Add some information to the file.
        Dim info As Byte()
        If intL = 0 Then
            strM = "There is a problem with this data:" & ChrW(10)
            strM = strM & "tblWorddocs: " & strF & ChrW(10)
            strM = strM & "Please contact your StudyDoc system administrator."
            info = New UTF8Encoding(True).GetBytes(strM)
            strpathT = Replace(strpathT, ".XML", ".TXT", 1, -1, CompareMethod.Text)
        Else
            ' Add some information to the file.
            info = New UTF8Encoding(True).GetBytes(strBuild.ToString())
        End If

        fs = File.Create(strpathT)
        fs.Close()
        fs = File.OpenWrite(strpathT)

        fs.Write(info, 0, info.Length)
        fs.Close()

        CreatexmlHome = strpathT

end1:

    End Function

    Function IsUserIDStuff(ByVal str1 As String) As Boolean

        IsUserIDStuff = False
        Select Case str1
            Case "CHARUSERIDINIT"
                IsUserIDStuff = True
            Case "CHARUSERNAMEINIT"
                IsUserIDStuff = True
            Case "DTINIT"
                IsUserIDStuff = True
            Case "CHARUSERIDMOD"
                IsUserIDStuff = True
            Case "CHARUSERNAMEMOD"
                IsUserIDStuff = True
            Case "DTMOD"
                IsUserIDStuff = True

        End Select


    End Function


    Function GetNewTempFileReport(boolClearDirs As Boolean) As String

        Dim boolE As Boolean
        Dim Count1 As Integer
        Dim strPathT1 As String
        Dim strPathT2 As String
        Dim strPathT3 As String
        Dim strPathT4 As String
        Dim strPathT5 As String

        Dim boolT1 As Boolean
        Dim boolT2 As Boolean
        Dim boolT3 As Boolean
        Dim boolT4 As Boolean
        Dim boolT5 As Boolean

        'clear all non-open files in /Temp
        If boolClearDirs Then
            Call ClearTemp()
        End If


        boolE = True
        Count1 = 0
        'strPathT1 = ""
        'strPathT2 = ""
        boolT1 = False
        boolT2 = False
        boolT3 = False
        boolT4 = False
        boolT5 = False

        'check to see if directory exists
        strPathT1 = "C:\Labintegrity\StudyDoc\TempReport\"
        If Directory.Exists(strPathT1) Then
        Else
            Directory.CreateDirectory(strPathT1)
        End If

        Do Until boolE = False
            Count1 = Count1 + 1
            strPathT1 = "C:\Labintegrity\StudyDoc\TempReport\TempReport" & Format(Count1, "00000") & ".xml"
            If File.Exists(strPathT1) Then
                boolT1 = False
            Else
                boolT1 = True
            End If

            strPathT2 = "C:\Labintegrity\StudyDoc\TempReport\TempReport" & Format(Count1, "00000") & ".docx"
            If File.Exists(strPathT2) Then
                boolT2 = False
            Else
                boolT2 = True
            End If

            strPathT3 = "C:\Labintegrity\StudyDoc\TempReport\TempReport" & Format(Count1, "00000") & ".pdf"
            If File.Exists(strPathT3) Then
                boolT3 = False
            Else
                boolT3 = True
            End If

            strPathT4 = "C:\Labintegrity\StudyDoc\TempReport\TempReport" & Format(Count1, "00000") & ".docm"
            If File.Exists(strPathT4) Then
                boolT4 = False
            Else
                boolT4 = True
            End If

            strPathT5 = "C:\Labintegrity\StudyDoc\TempReport\TempReport" & Format(Count1, "00000") & ".doc"
            If File.Exists(strPathT5) Then
                boolT5 = False
            Else
                boolT5 = True
            End If

            If boolT1 And boolT2 And boolT3 And boolT4 And boolT5 Then
                Exit Do
            End If
        Loop

        GetNewTempFileReport = strPathT1

    End Function

    Function GetNewTempFile(boolClearDirs As Boolean) As String

        Dim boolE As Boolean
        Dim Count1 As Integer
        Dim strPathT1 As String
        Dim strPathT2 As String
        Dim strPathT3 As String
        Dim strPathT4 As String
        Dim strPathT5 As String

        Dim boolT1 As Boolean
        Dim boolT2 As Boolean
        Dim boolT3 As Boolean
        Dim boolT4 As Boolean
        Dim boolT5 As Boolean


        boolE = True
        Count1 = 0
        strPathT1 = ""
        strPathT2 = ""
        boolT1 = False
        boolT2 = False
        boolT3 = False
        boolT4 = False
        boolT5 = False

        'clear all non-open files in /Temp
        If boolClearDirs Then
            Call ClearTemp()
        End If


        Do Until boolE = False
            Count1 = Count1 + 1
            strPathT1 = "C:\Labintegrity\StudyDoc\Temp\Temp" & Format(Count1, "00000") & ".xml"
            If File.Exists(strPathT1) Then
                boolT1 = False
            Else
                boolT1 = True
            End If

            strPathT2 = "C:\Labintegrity\StudyDoc\Temp\Temp" & Format(Count1, "00000") & ".docx"
            If File.Exists(strPathT2) Then
                boolT2 = False
            Else
                boolT2 = True
            End If

            strPathT3 = "C:\Labintegrity\StudyDoc\Temp\Temp" & Format(Count1, "00000") & ".pdf"
            If File.Exists(strPathT3) Then
                boolT3 = False
            Else
                boolT3 = True
            End If

            strPathT4 = "C:\Labintegrity\StudyDoc\Temp\Temp" & Format(Count1, "00000") & ".docm"
            If File.Exists(strPathT4) Then
                boolT4 = False
            Else
                boolT4 = True
            End If

            strPathT5 = "C:\Labintegrity\StudyDoc\Temp\Temp" & Format(Count1, "00000") & ".doc"
            If File.Exists(strPathT5) Then
                boolT5 = False
            Else
                boolT5 = True
            End If

            If boolT1 And boolT2 And boolT3 And boolT4 And boolT5 Then
                Exit Do
            End If
        Loop
        GetNewTempFile = strPathT1

    End Function

    Function Decode(ByVal input As String, ByVal boolAscii As Boolean)

        Dim str1 As String
        Dim var1, var2, var3
        Dim Count1 As Short
        Dim int1 As Short
        Dim varP, varPA
        Dim int2 As Short
        Dim int3 As Short
        Dim intS As Short
        Dim intE As Short
        Dim intE1 As Short

        If IsDBNull(input) Then
            Return ""
            Exit Function
        ElseIf Len(input) = 0 Then
            Return ""
            Exit Function
        End If

        If boolAscii Then 'return compiled ascii string
            'decypher ascii code
            varPA = input
            int1 = Len(varPA)
            var1 = ""
            int2 = 1
            int3 = 0
            intS = 1
            intE = 1
            intE1 = 1
            var3 = varPA
            For Count1 = 1 To int1
                intE1 = InStr(intS, varPA, " ", CompareMethod.Text)
                intE = intE1 - intS
                If intE1 = 0 Then
                    Exit For
                End If
                var2 = Mid(varPA, intS, intE)
                var1 = var1 & ChrW(var2)
                intS = intE1 + 1
            Next
            Return var1
        Else 'return de-compile ascii string
            int1 = Len(input)
            varPA = ""
            For Count1 = 1 To int1
                var1 = Mid(input, Count1, 1)
                varPA = varPA & AscW(var1) & " "
            Next
            Return varPA 'keep the space on the end
        End If

    End Function

    Private Function FixIt(ByVal v As Integer) As Integer
        'wraps a number to 0-255
        Do While v < 0 Or v > 255
            If v < 0 Then v += 255
            If v > 255 Then v -= 255
        Loop
        Return v
    End Function

    Function FindReportTypeID(ByRef tbl As System.Data.DataTable, ByRef str1 As String)

        Dim strF As String
        Dim Count1 As Short
        Dim ct1 As Short
        Dim var1
        Dim dv As System.Data.DataView
        Dim drows() As DataRow

        'dv = tbl.defaultview
        strF = "charReportType = '" & str1 & "'"
        drows = tbl.Select(strF)
        'dv.RowFilter = strF
        ct1 = drows.Length
        If ct1 = 0 Then
            FindReportTypeID = 0
        Else
            FindReportTypeID = dv.Item(0).Item("id_tblConfigReportType")
        End If

    End Function

    Function GetCPName(ByRef id As Int64, ByRef tbl As System.Data.DataTable)
        Dim dr() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim int1 As Short
        Dim ct1 As Short
        Dim Count1 As Short
        Dim Count2 As Short

        strF = "id_tblContributingPersonnel = " & id
        dr = tbl.Select(strF)

        str1 = ""
        'evaluate Prefix
        str2 = NZ(dr(0).Item("charCPPrefix"), "")
        If Len(str2) = 0 Then
        Else
            str1 = str1 & str2 & " "
        End If
        'evaluate Name
        str2 = NZ(dr(0).Item("charCPName"), "")
        If Len(str2) = 0 Then
        Else
            str1 = str1 & str2 & " "
        End If
        'evaluate Suffix
        str2 = NZ(dr(0).Item("charCPSuffix"), "")
        If Len(str2) = 0 Then
        Else
            str1 = str1 & str2 & " "
        End If
        'evaluate degree
        str2 = NZ(dr(0).Item("charCPDegree"), "")
        If Len(str2) = 0 Then
        Else
            str1 = str1 & ", " & str2
        End If

        GetCPName = str1


    End Function

    Function GetAddressTitle(ByRef charNickName As String, ByRef tbl As System.Data.DataTable) As String

        Dim dr() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim strS As String
        Dim int1 As Short
        Dim ct1 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim tblNick As System.Data.DataTable
        Dim rowsNick() As DataRow
        Dim var1
        Dim strLabel As String

        str2 = "charNickName = '" & NZ(charNickName, "[NA]") & "'"
        tblNick = tblCorporateNickNames
        rowsNick = tblNick.Select(str2)
        If rowsNick.Length = 0 Then
            var1 = -1 ' "[NA]"
        Else
            var1 = rowsNick(0).Item("id_tblCorporateNickNames")
        End If
        strF = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & True ' & " AND boolInclude = " & True


        GetAddressTitle = "[NA]"
        str1 = ""
        'strF = "charNickName = '" & charNickName & "' AND boolIncludeInTitle = " & True & " AND boolInclude = " & True
        strS = "id_tblAddressLabels ASC"
        dr = tbl.Select(strF, strS)
        int1 = dr.Length
        If int1 = 0 Then
        Else
            For Count1 = 0 To int1 - 1
                strLabel = dr(Count1).Item("CHARADDRESSLABEL")
                If Count1 = int1 - 1 Then
                    str1 = str1 & dr(Count1).Item("charValue")
                Else
                    Select Case strLabel
                        Case "City"
                            str1 = str1 & dr(Count1).Item("charValue") & ", "
                        Case "State/Province"
                            str1 = str1 & dr(Count1).Item("charValue") & " "
                        Case "Postal Code"
                            str1 = str1 & dr(Count1).Item("charValue") & Chr(10)
                        Case Else
                            str1 = str1 & dr(Count1).Item("charValue") & Chr(10)
                    End Select

                End If
            Next
            GetAddressTitle = str1
        End If


    End Function

    Function GetAddress(ByRef charNickName As String) As String

        Dim dr() As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim int1 As Short
        Dim ct1 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim tblNick As System.Data.DataTable
        Dim rowsNick() As DataRow
        Dim var1
        Dim intRows As Short
        Dim rows() As DataRow
        Dim tbl As System.Data.DataTable
        Dim strS As String

        str2 = "charNickName = '" & charNickName & "'"
        tblNick = tblCorporateNickNames
        rowsNick = tblNick.Select(str2)
        var1 = rowsNick(0).Item("id_tblCorporateNickNames")
        tbl = tblCorporateAddresses
        strF = "id_tblCorporateNickNames = " & var1
        strS = "ID_TBLADDRESSLABELS ASC"
        rows = tbl.Select(strF)
        intRows = rows.Length

        GetAddress = ""
        str1 = ""
        For Count1 = 0 To intRows - 1
            int1 = rows(Count1).Item("ID_TBLADDRESSLABELS")
            Select Case int1
                Case 1 'company
                    var1 = NZ(rows(Count1).Item("CHARVALUE"), "")
                    If Len(var1) = 0 Then
                    Else
                        str1 = var1
                    End If
                Case 2, 3, 4, 5 'addresses
                    var1 = NZ(rows(Count1).Item("CHARVALUE"), "")
                    If Len(var1) = 0 Then
                    Else
                        str1 = str1 & ChrW(11) & var1
                    End If
                Case 6 'city  20190111 LEE: change logic for city/state/postal.
                    '20190213 LEE: Aack! Logic is still wrong. city/state/postal must be inline
                    var1 = NZ(rows(Count1).Item("CHARVALUE"), "")
                    If Len(var1) = 0 Then
                    Else
                        str1 = str1 & ChrW(11) & var1
                    End If
                    'str1 = str1 & ChrW(11) & var1
                Case 7 'state/province
                    var1 = NZ(rows(Count1).Item("CHARVALUE"), "")
                    If Len(var1) = 0 Then
                    Else
                        str1 = str1 & ", " & var1
                    End If
                    'str1 = str1 & ", " & var1
                Case 8 'postal code
                    var1 = NZ(rows(Count1).Item("CHARVALUE"), "")
                    If Len(var1) = 0 Then
                    Else
                        str1 = str1 & " " & var1
                    End If
                    'str1 = str1 & " " & var1
                Case 9 'country
                    var1 = NZ(rows(Count1).Item("CHARVALUE"), "")
                    If Len(var1) = 0 Then
                    Else
                        str1 = str1 & ChrW(11) & var1
                    End If

            End Select

        Next

        GetAddress = str1

        'Select Case intCountry
        '    Case 1 'USA
        '        'get Name
        '        int1 = 1
        '        'strF = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = " & True & " AND boolInclude = " & True
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        ct1 = dr.Length
        '        For Count1 = 0 To ct1 - 1
        '            str2 = NZ(dr(Count1).Item("charValue"), "[NA]") & Chr(10)
        '            str1 = str1 & str2
        '        Next

        '        'get address1
        '        int1 = 2
        '        'strF = "charNickName = '" & charNickName & "' AND id_tblAddressLabels = " & int1 & " AND boolInclude = " & True
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        ct1 = dr.Length
        '        For Count1 = 0 To ct1 - 1
        '            str2 = NZ(dr(Count1).Item("charValue"), "[NA]") & Chr(10)
        '            str1 = str1 & str2
        '        Next

        '        'get address2
        '        int1 = 3
        '        'strF = "charNickName = '" & charNickName & "' AND id_tblAddressLabels = " & int1 & " AND boolInclude = " & True
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        ct1 = dr.Length
        '        For Count1 = 0 To ct1 - 1
        '            str2 = NZ(dr(Count1).Item("charValue"), "") & Chr(10)
        '            If Len(str2) = 0 Then
        '                Exit For
        '            End If
        '            str1 = str1 & str2
        '        Next

        '        'get address3
        '        int1 = 4
        '        'strF = "charNickName = '" & charNickName & "' AND id_tblAddressLabels = " & int1 & " AND boolInclude = " & True
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        ct1 = dr.Length
        '        For Count1 = 0 To ct1 - 1
        '            str2 = NZ(dr(Count1).Item("charValue"), "") & Chr(10)
        '            If Len(str2) = 0 Then
        '                Exit For
        '            End If
        '            str1 = str1 & str2
        '        Next

        '        'get address4
        '        int1 = 5
        '        'strF = "charNickName = '" & charNickName & "' AND id_tblAddressLabels = " & int1 & " AND boolInclude = " & True
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        ct1 = dr.Length
        '        For Count1 = 0 To ct1 - 1
        '            str2 = NZ(dr(Count1).Item("charValue"), "") & Chr(10)
        '            If Len(str2) = 0 Then
        '                Exit For
        '            End If
        '            str1 = str1 & str2
        '        Next

        '        'get City
        '        int1 = 6
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        'strF = "charNickName = '" & charNickName & "' AND id_tblAddressLabels = " & int1 & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        str2 = NZ(dr(0).Item("charValue"), "[NA]")
        '        str1 = str1 & str2 & ", "

        '        'get State
        '        int1 = 7
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        'strF = "charNickName = '" & charNickName & "' AND id_tblAddressLabels = " & int1 & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        str2 = NZ(dr(0).Item("charValue"), "[NA]")
        '        str1 = str1 & str2 & " "

        '        'get postal colde
        '        int1 = 8
        '        strF = "id_tblCorporateNickNames = " & var1 & " AND id_tblAddressLabels = " & int1 ' & " AND boolInclude = " & True
        '        'strF = "charNickName = '" & charNickName & "' AND id_tblAddressLabels = " & int1 & " AND boolInclude = " & True
        '        dr = tbl.Select(strF, "id_tblAddressLabels ASC")
        '        str2 = NZ(dr(0).Item("charValue"), "[NA]")
        '        str1 = str1 & str2

        '    Case 2 'Canada

        'End Select

        'GetAddress = str1

    End Function

    Function IsEven(ByVal num1) As Boolean
        Dim var1, var2, var3
        'works only for integers

        IsEven = False
        If Len(NZ(num1, "")) = 0 Then
            IsEven = False
        End If

        var1 = num1
        var2 = num1 / 2
        var3 = CInt(var2)

        If var2 = var3 Then
            IsEven = True
        Else
            IsEven = False
        End If

    End Function

    Function Pause(ByVal p As Single)

        'p is in seconds

        System.Threading.Thread.Sleep(p * 1000)


    End Function


    Function ReturnCalibrStds(ByVal strAnal As String, ByVal intGroup As Short, ByVal boolColumn As Boolean) As String

        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim int1 As Integer
        Dim Count1 As Integer
        Dim str1 As String
        Dim var1

        '20181129 LEE:
        'get calibration stds from tblCalStdGroupsAcc and intgroup
        If intGroup = -1 Then
            'find intgroup from tblAnalyteGroups
            strF = "ANALYTEDESCRIPTION_C = " & strAnal
            Dim rows1() As DataRow = tblAnalyteGroups.Select(strF)
            If rows1.Length = 0 Then
            Else
                intGroup = rows1(0).Item("INTGROUP")
            End If
        End If

        ReturnCalibrStds = "[NONE]"

        If intGroup = -1 Then
            tbl1 = tblBCStds
            strF = "ANALYTEDESCRIPTION = '" & CleanText(strAnal) & "'"
            strS = "CONCENTRATION ASC"
            rows = tbl1.Select(strF, strS)
            int1 = rows.Length
            If int1 = 0 Then
                GoTo end1
            End If
            var1 = rows(Count1).Item("CONCENTRATION")
            str1 = var1
            For Count1 = 1 To int1 - 2
                var1 = rows(Count1).Item("CONCENTRATION")
                str1 = str1 & ", " & var1
            Next
            var1 = rows(int1 - 1).Item("CONCENTRATION")
            str1 = str1 & ", and " & var1
        Else
            strF = "INTGROUP = " & intGroup
            Dim rows1() As DataRow = tblCalStdGroupsAcc.Select(strF, "CONCENTRATION ASC")
            int1 = rows1.Length
            If int1 = 0 Then
                GoTo end1
            End If
            If boolColumn Then
                For Count1 = 0 To rows1.Length - 1
                    If Count1 = 0 Then
                        str1 = rows1(Count1).Item("CONCENTRATION")
                    Else
                        str1 = str1 & ChrW(11) & rows1(Count1).Item("CONCENTRATION")
                    End If
                Next
            Else
                For Count1 = 0 To rows1.Length - 2
                    If Count1 = 0 Then
                        str1 = rows1(Count1).Item("CONCENTRATION")
                    Else
                        str1 = str1 & ", " & rows1(Count1).Item("CONCENTRATION")
                    End If
                Next
                var1 = rows1(rows1.Length - 1).Item("CONCENTRATION")
                str1 = str1 & ", and " & var1
            End If

        End If

        ReturnCalibrStds = str1

end1:
    End Function


    Function ReturnQCStds(ByVal strAnal As String, ByVal boolColumn As Boolean) As String

        Dim tbl1 As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim int1 As Integer
        Dim Count1 As Integer
        Dim str1 As String
        Dim var1

        'if study assays are messed up, levels may be different
        'Use Concentrations instead

        ReturnQCStds = "[NONE]"
        tbl1 = tblQCStds

        '20190206 LEE
        strF = "ANALYTEDESCRIPTION = '" & CleanText(strAnal) & "'"
        strS = "CONCENTRATION ASC"
        Dim dv As DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)

        Dim tbl2 As DataTable = dv.ToTable("a", True, "CONCENTRATION")

        rows = tbl1.Select(strF, strS)

        int1 = tbl2.Rows.Count
        If int1 = 0 Then
            GoTo end1
        End If
        Count1 = 0
        var1 = tbl2.Rows(Count1).Item("CONCENTRATION")
        str1 = var1

        If boolColumn Then
            For Count1 = 1 To int1 - 1
                var1 = tbl2.Rows(Count1).Item("CONCENTRATION")
                str1 = str1 & ChrW(11) & var1
            Next
        Else
            For Count1 = 1 To int1 - 2
                var1 = tbl2.Rows(Count1).Item("CONCENTRATION")
                str1 = str1 & ", " & var1
            Next
            var1 = tbl2.Rows(int1 - 1).Item("CONCENTRATION")
            str1 = str1 & ", and " & var1
        End If

      

        'int1 = rows.Length
        'If int1 = 0 Then
        '    GoTo end1
        'End If
        'var1 = rows(Count1).Item("CONCENTRATION")
        'str1 = var1
        'For Count1 = 1 To int1 - 2
        '    var1 = rows(Count1).Item("CONCENTRATION")
        '    str1 = str1 & ", " & var1
        'Next
        'var1 = rows(int1 - 1).Item("CONCENTRATION")
        'str1 = str1 & ", and " & var1

        ReturnQCStds = str1

end1:
    End Function


    Function UseAnalyteByTable(ByVal strA As String, boolQC As Boolean, boolRegr As Boolean) As Boolean

        UseAnalyteByTable = True

        Dim intTableID As Short


        If boolRegr Then
            intTableID = 2
        Else
            If boolQC Then
                intTableID = 4
            Else
                intTableID = 3
            End If
        End If

        Dim Count1 As Short
        Dim dv As System.Data.DataView = frmH.dgvReportTableConfiguration.DataSource
        Dim intRows As Short = dv.Count
        Dim tbl As System.Data.DataTable = dv.ToTable
        Dim rows() As DataRow
        Dim strF As String

        strF = "[" & strA & "] = " & True & " AND BOOLINCLUDE = " & True & " AND ID_TBLCONFIGREPORTTABLES = " & intTableID

        Try

            rows = tbl.Select(strF)

            If rows.Length = 0 Then
                UseAnalyteByTable = False
            Else
                UseAnalyteByTable = True
            End If

        Catch ex As Exception
            UseAnalyteByTable = False
        End Try



    End Function

    Function UseAnalyte(ByVal strA As String) As Boolean

        'This function returns true if Analyte is used in *any one* of the Included Tables in the Study
        UseAnalyte = True

        Dim Count1 As Short
        Dim dv As System.Data.DataView = frmH.dgvReportTableConfiguration.DataSource
        Dim intRows As Short = dv.Count
        Dim tbl As System.Data.DataTable = dv.ToTable
        Dim rows() As DataRow
        Dim strF As String

        strF = "[" & strA & "] = " & True & " AND BOOLINCLUDE = " & True

        Try

            rows = tbl.Select(strF)

            If rows.Length = 0 Then
                UseAnalyte = False
            Else
                UseAnalyte = True
            End If

        Catch ex As Exception
            UseAnalyte = False
        End Try



    End Function

    Function NZ(ByVal x, ByVal y)
        '
        If IsDBNull(x) Then
            NZ = y
        ElseIf Len(x) = 0 Then
            NZ = y
        Else
            NZ = x
        End If

    End Function

    Function GetWStudyID(ByVal id As Int64)

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String

        strF = "ID_TBLSTUDIES = " & id
        tbl = tblStudies
        rows = tbl.Select(strF)

        GetWStudyID = rows(0).Item("INT_WATSONSTUDYID")


    End Function

    Function IsInt(ByVal val)
        Dim var1, var2
        If IsNumeric(val) Then
        Else
            IsInt = False
            Exit Function
        End If
        var1 = CInt(val)
        If var1 = val Then
            IsInt = True
        Else
            IsInt = False
        End If
    End Function

    Function FindRowDVNumByCol(ByVal numSearch, ByVal dv, ByVal colName) 'finds row in a datatable
        Dim Count1 As Integer
        Dim rw As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim var1
        Dim int1 As Integer

        'dt = dg.DataSource
        'var1 = dt.name
        FindRowDVNumByCol = -1
        Count1 = -1
        int1 = dv.count
        For Count1 = 0 To int1 - 1
            var1 = NZ(dv.Item(Count1).Item(colName), 0)
            If var1 = numSearch Then
                FindRowDVNumByCol = Count1
                Exit For
            End If
        Next
        'dv = Nothing

    End Function

    Function FindRowDVByCol(ByVal strSearch, ByVal dv, ByVal colName) As Int16 'finds row in a datatable

        FindRowDVByCol = -1

        Dim Count1 As Integer
        Dim rw As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim var1
        Dim int1 As Integer

        Try
            'dt = dg.DataSource
            'var1 = dt.name
            FindRowDVByCol = -1
            Count1 = -1
            int1 = dv.count
            For Count1 = 0 To int1 - 1
                str1 = NZ(dv.Item(Count1).Item(colName), "")
                If StrComp(str1, strSearch, CompareMethod.Text) = 0 Then
                    FindRowDVByCol = Count1
                    Exit For
                End If
            Next
            'dv = Nothing
        Catch ex As Exception

        End Try


    End Function

    Function FindRowDV(ByVal strSearch As String, ByVal dv As System.Data.DataView) As Int16 'finds row in a datatable

        FindRowDV = -1

        Try
            Dim Count1 As Integer
            Dim rw As DataRow
            Dim str1 As String
            Dim str2 As String
            Dim var1
            Dim int1 As Integer

            'dt = dg.DataSource
            'var1 = dt.name
            FindRowDV = -1
            Count1 = -1

            '****
            int1 = dv.Count

            For Count1 = 0 To int1 - 1
                str1 = NZ(dv.Item(Count1).Item(0), "")
                If StrComp(str1, strSearch, CompareMethod.Text) = 0 Then
                    FindRowDV = Count1
                    Exit For
                End If
            Next
            'dv = Nothing
            '***
        Catch ex As Exception

        End Try



    End Function

    Function FindRow(ByVal strSearch, ByVal dt, ByVal strCol) As Int16 'finds row in a datatable

        Dim Count1 As Integer
        Dim rw As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim var1

        'dt = dg.DataSource
        'var1 = dt.name
        FindRow = -1
        Count1 = -1
        For Each rw In dt.Rows
            Count1 = Count1 + 1
            str1 = NZ(rw.Item(strCol), "")
            If StrComp(str1, strSearch, CompareMethod.Text) = 0 Then
                FindRow = Count1
                Exit For
            End If
        Next
        'dt = Nothing

    End Function

    Function GetMax(ByRef arr1, ByVal ct1)

        Dim Count1 As Int64
        Dim Count2 As Int64
        Dim var1, var2

        GetMax = 0

        Try
            'ignore blanks or null
            If ct1 = 0 Then
                GetMax = 0
            Else
                'GetMax = CDec(arr1(1))
                Count1 = 1
                Do Until Len(NZ(arr1(Count1), "")) <> 0
                    Count1 = Count1 + 1
                Loop
                var1 = arr1(Count1) 'debug
                GetMax = arr1(Count1)
                If ct1 < 2 Then
                Else
                    For Count2 = Count1 + 1 To ct1
                        var1 = NZ(arr1(Count2), "")
                        If Len(var1) = 0 Then 'ignore
                        Else
                            'var1 = CDec(var1)
                            var1 = var1
                            If var1 > GetMax Then
                                GetMax = var1
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            var2 = var1
        End Try



    End Function

    Function GetMin(ByRef arr1, ByVal ct1)

        Dim Count1 As Int64
        Dim Count2 As Int64
        Dim var1, var2, var3

        GetMin = 0

        Try
            'ignore blanks or null
            If ct1 = 0 Then
                GetMin = 0
            Else
                'GetMin = CDec(arr1(1))
                Count1 = 1
                var3 = arr1(Count1) 'debug
                Do Until Len(NZ(arr1(Count1), "")) <> 0
                    var3 = arr1(Count1) 'debug
                    Count1 = Count1 + 1
                Loop
                var3 = arr1(Count1) 'debug
                GetMin = arr1(Count1)
                If ct1 < 2 Then
                Else
                    For Count2 = Count1 + 1 To ct1
                        var3 = arr1(Count2) 'debug
                        var1 = NZ(arr1(Count2), "")
                        '''''''''''console.writeline(var1)
                        If Len(var1) = 0 Then 'ignore
                        Else
                            'var1 = CDec(var1)
                            'var1 = var1
                            If var1 < GetMin Then
                                GetMin = var1
                            End If
                        End If
                    Next
                    '''''''''''console.writeline("End")
                End If
            End If
        Catch ex As Exception
            var2 = var1

        End Try

    End Function

    Function AllCaps(ByVal s As String)
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim Count1 As Int64


        AllCaps = Trim(s)

        If Len(AllCaps) = 0 Then
            Exit Function
        End If

        AllCaps = UCase(AllCaps)

        Exit Function

        int1 = Len(s)

        For Count1 = 1 To int1
            str2 = Mid(s, Count1, 1)
            int3 = Asc(str2)
            If int3 < 123 And int3 > 96 Then 'needs capitalization
                int4 = int3 - 32
                AllCaps = Replace(s, Chr(int3), Chr(int4), 1, 1, CompareMethod.Text)
                s = AllCaps
            End If
        Next

    End Function

    Function Capit(ByVal s As String)

        'Note: Capit capitalizes the first letter of each 's'
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short

        Capit = Trim(s)

        If Len(Capit) = 0 Then
            Exit Function
        End If

        'str1 = Left(s, 1)
        str1 = Mid(s, 1, 1)
        int1 = Asc(str1)
        If int1 < 123 And int1 > 96 Then 'needs capitalization
            int2 = int1 - 32
            Capit = Chr(int2) & Right(s, Len(s) - 1)
        End If

    End Function

    Function DisplaySigFig(ByVal x, ByVal SigFigs) As String

        'this function will format trailing zeros if needed
        Dim intLen As Short
        Dim intDif As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim num1 As Int32
        Dim num2 As Int32
        Dim Count1 As Short

        DisplaySigFig = x

        If InStr(1, CStr(x), ".", CompareMethod.Text) > 0 Then
            intLen = Len(CStr(x)) - 1
        Else
            intLen = Len(CStr(x))
        End If
        If intLen < SigFigs Then
            intDif = SigFigs - intLen
            Select Case intDif
                Case 1
                    str1 = "0.0"
                Case 2
                    str1 = "0.00"
                Case 3
                    str1 = "0.000"
                Case 4
                    str1 = "0.0000"
                Case 5
                    str1 = "0.00000"
                Case 6
                    str1 = "0.000000"
            End Select

            DisplaySigFig = Format(x, str1)
        End If

        Exit Function


    End Function

    Function StrGubbs(ByVal str1, ByVal int1)
        Dim Count1 As Short
        Dim str2 As String
        str2 = str1
        For Count1 = 1 To int1 - 1
            str2 = str2 & str1
        Next
        StrGubbs = str2
    End Function

    Function DisplayCommas(ByVal x)

        Dim str3 As String

        DisplayCommas = x

        Dim Count1 As Short
        Dim xx As Int32

        If x < 1000 Then
            DisplayCommas = x
        Else
            If gINTCOMMAFORMAT = 0 Or CDbl(x) < gINTCOMMAFORMAT Then
                DisplayCommas = x
            Else
                str3 = "#,###.##"
                If Len(str3) = 0 Then
                    DisplayCommas = x
                Else
                    DisplayCommas = Format(CDec(x), str3)
                End If
            End If
        End If

    End Function

    Function DisplayNum(ByVal x, ByVal sigfigs, ByVal boolIgnore)

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim Count1 As Short
        Dim FormatSF As String
        Dim intSF As Int32
        Dim var1
        Dim xD As Decimal

        'boolIgnore = true if boolLSigFigs is to be ignored and use decimals instead

        '-- Answer sometimes needs padding with 0s --
        'If InStr(FormatSF, ".") = 0 Then
        '    If Len(FormatSF) < intSF Then
        'FormatSF = Format(FormatSF, "##0." & String(intSF - Len(FormatSF), "0"))
        '    End If
        'End If

        'If intSF > 1 And Abs(FormatSF) < 1 Then
        '    Do Until Left(Right(FormatSF, intSF), 1) <> "0" And Left(Right(FormatSF, intSF), 1) <> "."
        '        FormatSF = FormatSF & "0"
        '    Loop
        'End If

        '-- Answer sometimes needs padding with 0s --
        '-- From http://www.vbforums.com/showthread.php?s=&threadid=269312
        '-- had to modify a bit to correct

        Dim Count11 As Short
        Dim xx As Int32

        'If boolLUseSigFigs Or boolIgnore Then
        If boolIgnore = False Then

            'FormatSF = CStr(x)
            'intSF = LSigFig
            intSF = sigfigs
            '20151218 LEE: if x < 0.00001, cstr(x) returns unwanted scientific notation
            'solution is to use decimal data type for values less than 1
            ' 'http://stackoverflow.com/questions/1603013/easier-way-to-prevent-numbers-from-showing-in-exponent-notation
            If x < 1 Then
                xD = x
                str1 = CStr(xD)
            Else
                str1 = CStr(x)
            End If

            FormatSF = str1
            If InStr(str1, ".", CompareMethod.Text) = 0 Then
                If Len(str1) < intSF Then
                    'FormatSF = Format(FormatSF, "##0." & string("0",intSF - Len(FormatSF)))
                    str2 = "0." & StrGubbs("0", intSF - Len(str1))
                    FormatSF = Format(CDec(x), str2)
                End If
            ElseIf InStr(str1, ".", CompareMethod.Text) > 0 And (Len(str1) - 1) < intSF Then
                str2 = StrGubbs("0", intSF - (Len(str1) - 1))
                FormatSF = str1 & str2
            Else
                FormatSF = str1
            End If

            x = CDec(x)
            If Math.Abs(CDec(x)) = 0 Then
                var1 = FormatSF 'debugging
            Else
                'var1 = Math.Abs(CDec(x)) 'DEBUGGING
                If intSF > 1 And Math.Abs(CDec(x)) < 1 Then
                    var1 = Left(Right(FormatSF, intSF), 1)
                    Do Until var1 <> "0" And var1 <> "."
                        FormatSF = FormatSF & "0"
                        var1 = Left(Right(FormatSF, intSF), 1)
                    Loop
                End If
            End If

            If gINTCOMMAFORMAT = 0 Or CDbl(FormatSF) < gINTCOMMAFORMAT Then
                DisplayNum = FormatSF
            Else
                str3 = ""
                str3 = "#,###"
                'Select Case gINTCOMMAFORMAT
                '    Case Is < 10000
                '        str3 = "0,000" ' & str2
                '    Case Is < 100000
                '        str3 = "0,000" ' & str2
                '    Case Is < 1000000
                '        str3 = "0,000" ' & str2
                '    Case Is < 10000000
                '        str3 = "0,000,000" ' & str2 '1,000,000
                '    Case Is < 100000000
                '        str3 = "0,000,000" ' & str2 '10,000,000
                '    Case Is < 1000000000
                '        str3 = "0,000,000" ' & str2 '100,000,000
                '    Case Is < 10000000000
                '        str3 = "0,000,000,000" ' & str2 '1,000,000,000
                'End Select
                If Len(str3) = 0 Then
                    DisplayNum = FormatSF
                Else
                    DisplayNum = Format(CDec(FormatSF), str3)
                End If

            End If

        Else

            'this option should return decimal rounded with trailing 0's
            intSF = sigfigs ' LDec
            str1 = CStr(x)
            FormatSF = str1
            If intSF = 0 Then
                FormatSF = Format(CDec(x), "0")
            Else
                'str2 = "0." & StrGubbs("0", intSF - 1)
                'str2 = "." & StrGubbs("0", intSF - 1)
                str2 = "." & StrGubbs("0", intSF)


                If gINTCOMMAFORMAT = 0 Or CDbl(FormatSF) < gINTCOMMAFORMAT Then
                    str3 = "0" & str2
                Else
                    str3 = "#,###" & str2
                    'Select Case gINTCOMMAFORMAT
                    '    Case gINTCOMMAFORMAT <= 3
                    '        str3 = "0" & str2
                    '    Case gINTCOMMAFORMAT = 4 '1,000
                    '        str3 = "0,000" & str2
                    '    Case gINTCOMMAFORMAT = 5 '10,000
                    '        str3 = "0,000" & str2
                    '    Case gINTCOMMAFORMAT = 6 '100,000
                    '        str3 = "0,000,000" & str2 '1,000,000
                    '    Case gINTCOMMAFORMAT = 7
                    '        str3 = "0,000,000" & str2 '10,000,000
                    '    Case gINTCOMMAFORMAT = 8
                    '        str3 = "0,000,000" & str2 '100,000,000
                    '    Case gINTCOMMAFORMAT = 9
                    '        str3 = "0,000,000,000" & str2 '1,000,000,000
                    'End Select
                End If

                FormatSF = Format(CDec(x), str3)
            End If

            'If Math.Abs(x) = 0 Then
            '    var1 = FormatSF 'debugging
            'Else
            '    If intSF > 1 And Math.Abs(x) < 1 Then
            '        var1 = Left(Right(FormatSF, intSF), 1)
            '        Do Until var1 <> "0" And var1 <> "."
            '            FormatSF = FormatSF & "0"
            '            var1 = Left(Right(FormatSF, intSF), 1)
            '        Loop
            '    End If
            'End If
            DisplayNum = FormatSF

        End If

        Exit Function


    End Function

    Function GetAppPath() As String
        Dim mstrPath As String
        'mstrPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules(0).FullyQualifiedName)        
        mstrPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()(0).FullyQualifiedName)
        'If Right(mstrPath, 1) <> "\" Then        
        If mstrPath.Substring(mstrPath.Length - 1, 1) <> "\" Then
            mstrPath += "\"
        End If

        Return mstrPath
    End Function

    Function LowerCase(ByVal s As String)

        'boolAC = true if from anticoagulant, species, or certain items
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Integer
        Dim int2 As Integer
        Dim a As String
        Dim b As String
        Dim boolCont As Boolean
        Dim Count1 As Short
        Dim intCt As Short
        Dim intWdCt As Short
        Dim arrWord()

        LowerCase = Trim(s)

        If Len(LowerCase) = 0 Then
            Exit Function
        End If

        LowerCase = LCase(LowerCase)

        Exit Function

        intCt = Len(s)
        int1 = 1
        ReDim arrWord(intCt)

        For Count1 = 1 To intCt
            str1 = Mid(s, Count1, 1)
            arrWord(Count1) = str1
        Next

        For Count1 = 1 To intCt
            s = arrWord(Count1)
            int1 = Asc(s)
            If int1 > 64 And int1 < 91 Then 'needs uncapitalization
                int2 = int1 + 32
                'UnCapit = Chr(int2) & Right(s, Len(s) - 1)
                arrWord(Count1) = Chr(int2)
            End If

        Next

        'build uncapit
        str1 = ""
        For Count1 = 1 To intCt
            str1 = str1 & arrWord(Count1)
        Next

        LowerCase = str1

    End Function


    Function CapitAllWords(ByVal s As String)

        CapitAllWords = Trim(s)
        If Len(CapitAllWords) = 0 Then
            Exit Function
        End If

        Dim Count1 As Short
        Dim intCt As Short = Len(s)
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim str5 As String = ""
        Dim str6 As String = ""
        Dim arrWord() As String
        ReDim arrWord(intCt)
        Dim intWdCt As Short = 1
        Dim int1 As Short = 1
        Dim var1
        Dim intASC As Int16
        Dim strP As String

        'determine how many words, spaces and hyphens are in s
        For Count1 = 1 To intCt
            Try
                str1 = Mid(s, Count1, 1)
            Catch ex As Exception
                var1 = ex.Message
            End Try
            Try
                intASC = AscW(str1)
                Select Case intASC
                    'Space: 32
                    'NBSP: 160
                    'Hyphen: 45
                    'NBH: 173, 2011, 8209
                    Case 32, 160, 45, 173, 2011, 8209

                        'record word
                        str4 = Mid(s, int1, Count1 - int1)
                        arrWord(intWdCt) = str4
                        intWdCt = intWdCt + 1

                        'record space or hyphen
                        arrWord(intWdCt) = str1
                        int1 = Count1 + 1
                        intWdCt = intWdCt + 1

                        'str4 = Mid(s, int1, Count1 - int1)
                        'Select Case intASC
                        '    Case 45, 173, 2011, 8209
                        '        str4 = ChrW(intASC) & str4
                        'End Select
                        'arrWord(intWdCt) = str4
                        'int1 = Count1 + 1
                        'intWdCt = intWdCt + 1
                End Select
            Catch ex As Exception
                var1 = ex.Message
            End Try
        Next
        If intWdCt = 1 Then
            arrWord(1) = s
        Else
            'get that last word
            arrWord(intWdCt) = Mid(s, int1, intCt - int1 + 1)
        End If

        'Space: 32
        'NBSP: 160
        'Hyphen: 45
        'NBH: 173, 2011, 8209
        If intWdCt = 1 Then
            str2 = Capit(s)
        Else
            str3 = ""
            For Count1 = 1 To intWdCt

                str2 = arrWord(Count1)
                If Len(str2) = 1 Then
                    intASC = AscW(str2)
                Else
                    intASC = -1
                End If
                str4 = UCase(str2)
                Select Case str4
                    Case "AT", "FOR"
                        str1 = str2
                    Case Else
                        Select Case intASC
                            Case 32, 160, 45, 173, 2011, 8209
                                str1 = str2
                            Case Else
                                str1 = Capit(str2)
                        End Select

                End Select

                If Count1 = 1 Then
                    str3 = str1
                Else
                    If IsNumeric(strP) Then
                        ''make sure next word doesn't start with hyphen
                        'str4 = Mid(str1, 1, 1)
                        Select Case intASC
                            Case 32
                                str3 = str3 & ChrW(160)
                            Case Else
                                str3 = str3 & str1
                        End Select

                    Else
                        str3 = str3 & str1
                    End If

                End If
                strP = str1
            Next
            str2 = str3
        End If

        CapitAllWords = str2

    End Function

    Function UnCapit(ByVal s As String, ByVal boolAC As Boolean)

        'boolAC = true if from anticoagulant, species, or certain items
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Integer
        Dim int2 As Integer
        Dim a As String
        Dim b As String
        Dim boolCont As Boolean
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim intCt As Short
        Dim intWdCt As Short
        Dim arrWord()

        UnCapit = Trim(s)

        If Len(UnCapit) = 0 Then
            Exit Function
        End If

        'UnCapit = LCase(UnCapit)

        'Exit Function

        intCt = Len(s)
        int1 = 1
        intWdCt = 1

        ReDim arrWord(100)
        If boolAC Then 'do this only for anticoagulant
            'first determine how many words are in s
            For Count1 = 1 To intCt
                str1 = Mid(s, Count1, 1)
                If StrComp(str1, " ", CompareMethod.Text) = 0 Then
                    arrWord(intWdCt) = Mid(s, int1, Count1 - int1)
                    int1 = Count1 + 1
                    If int1 > UBound(arrWord) Then
                        ReDim Preserve arrWord(UBound(arrWord) + 100)
                    End If
                    intWdCt = intWdCt + 1
                End If
            Next
            If intWdCt = 1 Then
                arrWord(1) = s
            Else
                'get that last word
                arrWord(intWdCt) = Mid(s, int1, intCt - int1 + 1)
            End If
        Else
            arrWord(intWdCt) = s

        End If


        For Count1 = 1 To intWdCt
            'if s=acronym, then ignore
            s = arrWord(Count1)
            boolCont = True

            Dim ts As String

            If Len(s) <= 1 Then 'don't evaluate further
                'boolCont = False

            Else

                '20170708 LEE: New logic. Current logic gives false positive for something like K3EDTA or NaSO4
                'instead, count numbers and capital letters
                'if numbers + capital letters > 1/2 of all letters, then is acronym
                Dim intCapit As Short = 0
                Dim intNum As Short = 0
                Dim intTot As Short = 0
                Dim numCrit As Single = 0.5

                For Count2 = 1 To Len(s)
                    a = Mid(s, Count2, 1)
                    int1 = Asc(a)

                    If (int1 > 64 And int1 < 91) Then
                        intCapit = intCapit + 1
                    ElseIf (int1 > 47 And int1 < 58) Then
                        intNum = intNum + 1
                    End If

                Next

                intTot = intCapit + intNum
                If intTot = 0 Then
                    'not an acronym
                    ts = LowerCase(s)
                Else
                    If intTot > Len(s) * numCrit Then
                        'is an acronym
                        'don't do anything
                        ts = s
                    Else
                        ts = LowerCase(s)
                    End If
                End If


                'a = Mid(s, 1, 1)
                'int1 = Asc(a)
                'If (int1 > 64 And int1 < 91) Or (int1 > 96 And int1 < 123) Then
                '    'check to see if next character is capital or numeric
                '    b = Mid(s, 2, 1)
                '    int1 = Asc(b)
                '    If (int1 > 64 And int1 < 91) Then 's is probably an acronym
                '        boolCont = False
                '    End If
                'Else 'acronym with a non-alpha character
                '    boolCont = False
                'End If
            End If

            arrWord(Count1) = ts

            'If boolCont Then
            '    'str1 = Left(s, 1)
            '    str1 = Mid(s, 1, 1)
            '    int1 = Asc(str1)
            '    If int1 > 64 And int1 < 91 Then 'needs uncapitalization
            '        int2 = int1 + 32
            '        'UnCapit = Chr(int2) & Right(s, Len(s) - 1)
            '        arrWord(Count1) = Chr(int2) & Mid(s, 2, Len(s) - 1)
            '    End If
            'End If
        Next

        'build uncapit
        str1 = ""
        For Count1 = 1 To intWdCt
            If Count1 = 1 Then
                str1 = arrWord(Count1)
            Else
                str1 = str1 & " " & arrWord(Count1)
            End If

        Next
        UnCapit = str1


    End Function

    Function FindFirstDigit(x) As Short

        FindFirstDigit = 0

        'To account for potential floating point issues
        'Correct NumDigitsAfterDecimal for first non-zero digit after decimal

        If x < 1 Then
            'find first digit
            Dim str1 As String = CStr(x)
            Dim str2 As String
            Dim Count1 As Short
            Dim intDec As Short
            Dim int1 As Short
            Dim intMoreDigits As Short = 0

            'e.g.: 0.957439
            '20170904 LEE: function fails if number is negative
            '   CInt(string) fails
            'add isnumeric evalution
            intDec = InStr(1, str1, ".", CompareMethod.Text)
            For Count1 = intDec + 1 To Len(str1)
                str2 = Mid(str1, Count1, 1)
                If IsNumeric(str2) Then
                    If CInt(str2) = 0 Then 'ignore
                    Else
                        FindFirstDigit = Count1 - 2
                        Exit For
                    End If
                Else 'ignore

                End If

            Next
        End If

    End Function

    Public Function RoundToDecimalRAFZ(ByVal Number As Object, Optional ByVal NumDigitsAfterDecimal As Integer = 0) As Decimal

        RoundToDecimalRAFZ = RoundToDecimal(Number, NumDigitsAfterDecimal)

    End Function

    Public Function RoundToDecimalA(ByVal Number As Object, Optional ByVal NumDigitsAfterDecimal As Integer = 0) As Decimal

        'http://anderly.com/2009/07/28/to-round-up-or-to-round-down-that-is-the-question/
        'https://support.microsoft.com/en-us/kb/196652

        'IEEE Standard 754 Section 4: round 5 to even. Microsoft development products (VBA, VB, .NET) math.round function uses this convention
        'Bankers Rounding: round to even. 
        'Excel ROUND function does round away from 0. So does JAVA. So does VB 'Format' function

        'To account for potential floating point issues
        'Correct NumDigitsAfterDecimal for first non-zero digit after decimal
        'Then add a number of exta digits
        Dim intExtraDigits As Short = 13 - NumDigitsAfterDecimal
        Dim int1 As Short = FindFirstDigit(Number)
        NumDigitsAfterDecimal = NumDigitsAfterDecimal + int1 + intExtraDigits

        Dim var1
        Try
            If gboolRoundFiveAway Then
                RoundToDecimalA = CDec(FormatNumber(Number, NumDigitsAfterDecimal)) ' this rounds 5 away from 0
            Else
                If (Number < Decimal.MaxValue) And (Number > Decimal.MinValue) Then  'Use Decimal values if we can: these avoid the issues rounding doubles.
                    RoundToDecimalA = CDec(Math.Round(Math.Round(CDec(Number), NumDigitsAfterDecimal + 5), NumDigitsAfterDecimal)) 'this rounds 5 even
                Else
                    RoundToDecimalA = CDec(Math.Round(Number, NumDigitsAfterDecimal)) 'this rounds 5 even  'Occasionally, might have issues rounding doubles
                End If
            End If
        Catch ex As Exception
            RoundToDecimalA = 0
            var1 = ex.Message
        End Try


    End Function

    Public Function RoundToDecimal(ByVal Number As Object, Optional ByVal NumDigitsAfterDecimal As Integer = 0) As Decimal

        'http://anderly.com/2009/07/28/to-round-up-or-to-round-down-that-is-the-question/
        'https://support.microsoft.com/en-us/kb/196652

        'IEEE Standard 754 Section 4: round 5 to even. Microsoft development products (VBA, VB, .NET) math.round function uses this convention
        'Bankers Rounding: round to even. 
        'Excel ROUND function does round away from 0. So does JAVA. So does VB 'Format' function

        Dim var1
        Try
            If gboolRoundFiveAway Then
                RoundToDecimal = CDbl(FormatNumber(Number, NumDigitsAfterDecimal)) ' this rounds 5 away from 0
            Else
                If (Number < Decimal.MaxValue) And (Number > Decimal.MinValue) Then  'Use Decimal values if we can: these avoid the issues rounding doubles.
                    RoundToDecimal = CDec(Math.Round(Math.Round(CDec(Number), NumDigitsAfterDecimal + 5), NumDigitsAfterDecimal)) 'this rounds 5 even
                Else
                    RoundToDecimal = CDec(Math.Round(Number, NumDigitsAfterDecimal)) 'this rounds 5 even  'Occasionally, might have issues rounding doubles
                End If
            End If
        Catch ex As Exception
            RoundToDecimal = 0
            var1 = ex.Message
        End Try


    End Function

    Function ANOVA_OneWay(ByVal tblAnova As System.Data.DataTable)

        Dim strF As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim var1, var2, var3, var4, var5
        Dim sumn
        Dim sumX
        Dim sumX2
        Dim sumC2
        Dim ssFactor 'Between
        Dim ssError 'Within
        Dim dfFactor 'Between
        Dim dfError 'Within
        Dim MSFactor 'Between
        Dim MSError 'Within
        Dim tblB As New System.Data.DataTable
        Dim dvAnova As System.Data.DataView
        Dim intA As Short
        Dim int1 As Short
        Dim N 'total number of replicates in dataset
        Dim ms 'average sample size
        Dim c 'number of runs
        Dim xbar 'mean of all values
        Dim BR
        Dim WR
        Dim rRet(1)

        BR = "NA"
        WR = "NA"

        'configure tblB
        Dim col1 As New DataColumn
        col1.ColumnName = "ColumnTotal"
        col1.DataType = System.Type.GetType("System.Decimal")
        tblB.Columns.Add(col1)
        Dim col2 As New DataColumn
        col2.ColumnName = "NumReps"
        col2.DataType = System.Type.GetType("System.Int16")
        tblB.Columns.Add(col2)

        For Count3 = 1 To 2



        Next

        'first get sum(X) and sum(X2)
        intRows = tblAnova.Rows.Count
        N = intRows
        sumX = 0
        sumX2 = 0
        '''''''''''''console.writeline("StartX")
        For Count1 = 0 To intRows - 1
            var1 = tblAnova.Rows.Item(Count1).Item("Conc")
            '''''''''''''console.writeline(var1)
            sumX = sumX + var1
            var2 = var1 ^ 2
            sumX2 = sumX2 + var2
        Next
        '''''''''''''console.writeline("EndX")
        xbar = SigFigOrDec(sumX / N, LSigFig, False)

        'fill tblB
        'find distinct groups
        dvAnova = tblAnova.DefaultView
        Dim tblT As System.Data.DataTable = dvAnova.ToTable("a", True, "Group")
        intA = tblT.Rows.Count
        c = intA
        sumC2 = 0
        sumn = 0
        For Count1 = 0 To c - 1
            var1 = tblT.Rows.Item(Count1).Item("Group")
            strF = "Group = " & var1
            dvAnova.RowFilter = ""
            dvAnova.RowFilter = strF
            int1 = dvAnova.Count
            var1 = 0
            var2 = 0
            var5 = 0
            For Count2 = 0 To int1 - 1
                var1 = dvAnova(Count2).Item("Conc")
                ''''''''''''''console.writeline(var1)
                var5 = var5 + var1
            Next
            var2 = var5 ^ 2
            var3 = var2 / int1
            sumC2 = sumC2 + var3

            'estimate ms
            var4 = int1 ^ 2
            sumn = sumn + var4
        Next

        ssFactor = sumC2 - ((sumX ^ 2) / N)
        ssError = sumX2 - sumC2
        dfFactor = intA - 1
        dfError = N - c
        If dfFactor = 0 Then

            rRet(0) = "NA"
            rRet(1) = "NA"

        Else
            MSFactor = ssFactor / dfFactor
            MSError = ssError / dfError

            'estimate ms
            var1 = sumn / N
            var2 = N - var1
            If (c - 1) = 0 Then
                var3 = 0
            Else
                var3 = var2 / (c - 1)
            End If

            ms = var3

            'between run precision
            If MSError > MSFactor Or xbar = 0 Or ms = 0 Then
                var1 = 0
            Else
                var1 = (((MSFactor - MSError) / ms) ^ 0.5) / xbar * 100
            End If
            BR = RoundToDecimal(var1, 1)
            rRet(0) = BR

            'within run precision
            If xbar = 0 Then
                var1 = 0
            Else
                var1 = ((MSError ^ 0.5) / xbar) * 100
            End If

            WR = RoundToDecimal(var1, 1)
            rRet(1) = WR

        End If


        'Return (rRet)
        ANOVA_OneWay = rRet

    End Function

    Function StdDevDRArea(ByVal r() As DataRow, ByVal strCol As String, ByVal boolAliq As Boolean, ByVal strAliq As String, ByVal boolSF As Boolean, ByVal boolUseIS As Boolean)

        '20160119 LEE:  Deprecated. This function not needed

        Dim i As Short
        Dim k As Short
        Dim avg As Object, SumSq As Object
        Dim var1, var2, var3
        Dim varIS, varAnal
        Dim num1 As Double

        StdDevDRArea = 0

        k = r.Length

        If k < 2 Then 'can't do sd on n=1
            StdDevDRArea = 0
            Exit Function
        End If

        avg = CDec(MeanDRArea(r, strCol, boolAliq, strAliq, boolSF, boolUseIS))
        'avg = CDec(MeanDR(r, strCol, boolAliq, strAliq, boolSF, boolUseIS))

        ''calculate own average
        'Dim Count1 As Short
        'num1 = 0
        'For Count1 = 0 To k - 1
        '    var1 = r(Count1)
        '    If boolLUseSigFigsArea Then
        '        If LboolWyethRoundingArea Then
        '            var1 = RoundToDecimalRAFZ(var1, 0)
        '        Else
        '            var1 = SigFigArea(var1, LSigFigArea, True, False)
        '        End If
        '    Else
        '        var1 = RoundToDecimalRAFZ(var1, LDecArea)
        '    End If
        '    num1 = num1 + var1
        'Next
        'avg = num1 / k

        If boolSF Then 'prepare data with appropriate sigfigs
            If boolAliq Then
                For i = 0 To k - 1
                    var1 = NZ(r(i).Item(strCol), 0)
                    If boolLUseSigFigsArea Then
                        If LboolWyethRoundingArea Then
                            var1 = RoundToDecimalRAFZ(var1, 0)
                        Else
                            var1 = SigFigArea(var1, LSigFigArea, True, False)
                        End If
                    Else
                        var1 = RoundToDecimalRAFZ(var1, LDecArea)
                    End If
                    var2 = NZ(r(i).Item(strAliq), 1)
                    var3 = var1 / var2 ' sigfigarea(RoundToDecimal(var1 / var2, 5), LSigFig, False)
                    If boolLUseSigFigsArea Then
                        If LboolWyethRoundingArea Then
                            var3 = RoundToDecimalRAFZ(var3, 0)
                        Else
                            var3 = SigFigArea(var3, LSigFigArea, True, False)
                        End If
                    Else
                        var3 = RoundToDecimalRAFZ(var3, LDecArea)
                    End If
                    SumSq = SumSq + (var3 - avg) ^ 2
                Next i
            Else
                For i = 0 To k - 1
                    'SumSq = SumSq + (arr(i) - avg) ^ 2
                    var1 = NZ(r(i).Item(strCol), 0)
                    If boolLUseSigFigsArea Then
                        If LboolWyethRoundingArea Then
                            var1 = RoundToDecimalRAFZ(var1, 0)
                        Else
                            var1 = SigFigArea(var1, LSigFigArea, True, False)
                        End If
                    Else
                        var1 = RoundToDecimalRAFZ(var1, LDecArea)
                    End If
                    var3 = var1 ' sigfigarea(RoundToDecimal(var1, 5), LSigFig, False)
                    SumSq = SumSq + (var3 - avg) ^ 2
                Next i
            End If
        Else
            If boolAliq Then
                For i = 0 To k - 1
                    var1 = NZ(r(i).Item(strCol), 0)
                    If boolLUseSigFigsArea Then
                        If LboolWyethRoundingArea Then
                            var1 = RoundToDecimalRAFZ(var1, 0)
                        Else
                            var1 = SigFigArea(var1, LSigFigArea, True, False)
                        End If
                    Else
                        var1 = RoundToDecimalRAFZ(var1, LDecArea)
                    End If
                    var2 = NZ(r(i).Item(strAliq), 1)
                    var3 = var1 / var2
                    If boolLUseSigFigsArea Then
                        If LboolWyethRoundingArea Then
                            var3 = RoundToDecimalRAFZ(var3, 0)
                        Else
                            var3 = SigFigArea(var3, LSigFigArea, True, False)
                        End If
                    Else
                        var3 = RoundToDecimalRAFZ(var3, LDecArea)
                    End If
                    SumSq = SumSq + (var3 - avg) ^ 2
                Next i
            Else
                For i = 0 To k - 1
                    'SumSq = SumSq + (arr(i) - avg) ^ 2
                    var1 = NZ(r(i).Item(strCol), 0)

                    Try
                        varAnal = NZ(r(i).Item("ANALYTEAREA"), 0)
                    Catch ex As Exception
                        varAnal = 0
                    End Try
                    'If boolLUseSigFigsArea Then
                    '    If LboolWyethRoundingArea Then
                    '        varAnal = RoundToDecimalRAFZ(varAnal, 0)
                    '    Else
                    '        varAnal = SigFigArea(varAnal, LSigFigArea, True, False)
                    '    End If
                    'Else
                    '    varAnal = RoundToDecimalRAFZ(varAnal, LDecArea)
                    'End If

                    Try
                        varIS = NZ(r(i).Item("INTERNALSTANDARDAREA"), 0)
                    Catch ex As Exception
                        varIS = 0
                    End Try
                    'If boolLUseSigFigsArea Then
                    '    If LboolWyethRoundingArea Then
                    '        varIS = RoundToDecimalRAFZ(varIS, 0)
                    '    Else
                    '        varIS = SigFigArea(varIS, LSigFigArea, True, False)
                    '    End If
                    'Else
                    '    varIS = RoundToDecimalRAFZ(varIS, LDecArea)
                    'End If

                    If boolRCPARatio Then
                        '20180719 LEE:
                        If StrComp(strCol, "INTERNALSTANDARDAREA", CompareMethod.Text) = 0 Then
                            var1 = varIS
                            '20180719 LEE:
                            'Here, use LSigFigArea, not LSigFigAreaRatio
                            If boolLUseSigFigsArea Then
                                If LboolWyethRoundingArea Then
                                    var1 = RoundToDecimalRAFZ(var1, 0)
                                Else
                                    var1 = SigFigArea(var1, LSigFigArea, True, False)
                                End If
                            Else
                                var1 = RoundToDecimalRAFZ(var1, LDecArea)
                            End If
                        Else
                            If varIS = 0 Then
                                var1 = 0
                            Else

                                'varAnal = SigFigArea(varAnal, LSigFigArea, True, False)
                                'varIS = SigFigArea(varIS, LSigFigArea, True, False)

                                var1 = varAnal / varIS ' RoundToDecimalRAFZ(varAnal / varIS, 5)
                                If boolLUseSigFigsAreaRatio Then
                                    If LboolWyethRoundingArea Then
                                        var1 = RoundToDecimalRAFZ(var1, 0)
                                    Else
                                        var1 = SigFigAreaRatio(var1, LSigFigAreaRatio, True, False)
                                    End If
                                Else
                                    var1 = RoundToDecimalRAFZ(var1, LDecAreaRatio)
                                End If
                            End If
                        End If
                        'var1 = varIS ' RoundToDecimalRAFZ(varIS, 0)
                        'If boolLUseSigFigsArea Then
                        '    If LboolWyethRoundingArea Then
                        '        var1 = RoundToDecimalRAFZ(var1, 0)
                        '    Else
                        '        var1 = SigFigArea(var1, LSigFigArea, True, False)
                        '    End If
                        'Else
                        '    var1 = RoundToDecimalRAFZ(var1, LDecArea)
                        'End If
                        'ElseIf boolRCPARatio Then
                        '    If varIS = 0 Then
                        '    Else
                        '        var1 = varAnal / varIS ' RoundToDecimalRAFZ(varAnal / varIS, 5)
                        '        If boolLUseSigFigsAreaRatio Then
                        '            If LboolWyethRoundingArea Then
                        '                var1 = RoundToDecimalRAFZ(var1, 0)
                        '            Else
                        '                var1 = SigFigAreaRatio(var1, LSigFigAreaRatio, True, False)
                        '            End If
                        '        Else
                        '            var1 = RoundToDecimalRAFZ(var1, 5)
                        '        End If
                        '    End If

                    ElseIf boolRCPA Then '20180719 LEE:

                        If boolUseIS Then
                            var1 = varIS ' RoundToDecimalRAFZ(varIS, 0)
                        Else
                            var1 = varAnal ' RoundToDecimalRAFZ(varAnal, 0)
                        End If
                        'var1 = varAnal ' RoundToDecimalRAFZ(varAnal, 0)
                        If boolLUseSigFigsArea Then
                            If LboolWyethRoundingArea Then
                                var1 = RoundToDecimalRAFZ(var1, 0)
                            Else
                                var1 = SigFigArea(var1, LSigFigArea, True, False)
                            End If
                        Else
                            var1 = RoundToDecimalRAFZ(var1, LDecArea)
                        End If
                        'ElseIf boolRCPA And boolUseIS = False Then
                        '    var1 = varAnal ' RoundToDecimalRAFZ(varAnal, 0)
                        'ElseIf boolRCPA And boolUseIS Then
                        '    var1 = varIS ' RoundToDecimalRAFZ(varIS, 0)
                    Else
                        If boolLUseSigFigsArea Then
                            If LboolWyethRoundingArea Then
                                var1 = RoundToDecimalRAFZ(var1, 0)
                            Else
                                var1 = SigFigArea(var1, LSigFigArea, True, False)
                            End If
                        Else
                            var1 = RoundToDecimalRAFZ(var1, LDecArea)
                        End If
                    End If

                    SumSq = SumSq + (var1 - avg) ^ 2
                Next i
            End If
        End If

        var1 = Math.Sqrt(SumSq / (k - 1))
        'If boolLUseSigFigsArea Then
        '    If LboolWyethRoundingArea Then
        '        var1 = RoundToDecimalRAFZ(var1, 0)
        '    Else
        '        var1 = SigFigArea(var1, LSigFigArea, True, False)
        '    End If
        'Else
        '    var1 = RoundToDecimalRAFZ(var1, LDecArea)
        'End If
        StdDevDRArea = var1


    End Function

    Function StdDevDR(ByVal r() As DataRow, ByVal strCol As String, ByVal boolAliq As Boolean, ByVal strAliq As String, ByVal boolSF As Boolean, ByVal boolUseIS As Boolean)

        Dim i As Short
        Dim k As Short
        Dim avg As Object, SumSq As Object
        Dim var1, var2, var3
        Dim varIS As Decimal
        Dim varAnal As Decimal
        Dim boolHasAnalyte As Boolean = False
        Dim sk As Short
        Dim d1 As Decimal
        Dim d2 As Decimal
        Dim d3 As Decimal

        StdDevDR = 0

        k = r.Length

        If k < 2 Then 'can't do sd on n=1
            StdDevDR = 0
            Exit Function
        End If

        avg = CDec(MeanDR(r, strCol, boolAliq, strAliq, boolSF, boolUseIS))

        If boolSF Then 'prepare data with appropriate sigfigs
            If boolAliq Then
                For i = 0 To k - 1
                    var1 = NZ(r(i).Item(strCol), 0)
                    var2 = NZ(r(i).Item(strAliq), 1)
                    d1 = var1
                    d2 = var2
                    d3 = SigFigOrDec(RoundToDecimalA(d1 / d2, LSigFig), LSigFig, False)
                    SumSq = SumSq + (d3 - avg) ^ 2
                Next i
            Else
                For i = 0 To k - 1
                    'SumSq = SumSq + (arr(i) - avg) ^ 2
                    var1 = NZ(r(i).Item(strCol), 0)
                    d1 = var1
                    d3 = SigFigOrDec(RoundToDecimalA(d1, LSigFig), LSigFig, False)
                    SumSq = SumSq + (d3 - avg) ^ 2
                Next i
            End If
        Else
            If boolAliq Then
                For i = 0 To k - 1
                    var1 = NZ(r(i).Item(strCol), 0)
                    var2 = NZ(r(i).Item(strAliq), 1)
                    d1 = var1
                    d2 = var2
                    d3 = d1 / d2
                    SumSq = SumSq + (d3 - avg) ^ 2
                Next i
            Else
                sk = 0
                For i = 0 To k - 1
                    'SumSq = SumSq + (arr(i) - avg) ^ 2
                    var1 = NZ(r(i).Item(strCol), 0)
                    If IsNumeric(var1) Then
                        var1 = CDec(var1)
                        sk = sk + 1
                        boolHasAnalyte = True
                        Try
                            varAnal = NZ(r(i).Item("ANALYTEAREA"), 0)
                        Catch ex As Exception
                            varAnal = 0
                            boolHasAnalyte = False
                        End Try
                        Try
                            varIS = NZ(r(i).Item("INTERNALSTANDARDAREA"), 0)
                        Catch ex As Exception
                            varIS = 0
                        End Try

                        '*****

                        If StrComp(strCol, "INTERNALSTANDARDAREA", CompareMethod.Text) = 0 Then
                            If boolLUseSigFigsArea Then
                                var1 = SigFigArea(RoundToDecimalA(varIS, LSigFigArea), LSigFigArea, True, False) 'special rounding incorporated
                            Else
                                var2 = Format(RoundToDecimalRAFZ(varIS, LSigFigArea), GetRegrDecStr(LSigFigArea))
                            End If
                        Else
                            If boolHasAnalyte Then
                                If boolRCPARatio Then

                                    If varIS = 0 Then
                                        var1 = 0
                                    Else
                                        If boolLUseSigFigsAreaRatio Then
                                            var1 = SigFigAreaRatio(RoundToDecimalA(varAnal / varIS, LSigFigAreaRatio), LSigFigAreaRatio, True, False) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(varAnal / varIS, LSigFigAreaRatio), GetRegrDecStr(LSigFigAreaRatio))
                                        End If
                                    End If
                                    'ElseIf boolRCPARatio Then
                                    '    If varIS = 0 Then
                                    '        var1 = 0
                                    '    Else
                                    '        var1 = varAnal / varIS ' RoundToDecimalRAFZ(varAnal / varIS, 5)
                                    '    End If
                                ElseIf boolRCPA Then
                                    'var1 = RoundToDecimalRAFZ(varAnal, 0)
                                    If boolLUseSigFigsArea Then
                                        var1 = SigFigArea(RoundToDecimalA(varAnal, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(varAnal, LSigFigArea), GetRegrDecStr(LSigFigArea))
                                    End If
                                    'ElseIf boolRCPA And boolUseIS Then
                                    '    'var1 = RoundToDecimalRAFZ(varIS, 0)
                                Else
                                    var1 = NZ(r(i).Item(strCol), 0)
                                    var2 = NZ(r(i).Item(strAliq), 1)
                                    d1 = var1
                                    d2 = var2
                                    d3 = d1 / d2 ' SigFigOrDec(RoundToDecimalA(var1 / var2, LSigFig), LSigFig, False)
                                    If boolLUseSigFigs Then
                                        var1 = SigFigOrDec(RoundToDecimalA(d3, LSigFig), LSigFig, False) 'special rounding incorporated
                                    Else
                                        var1 = Format(RoundToDecimalRAFZ(d3, LSigFig), GetRegrDecStr(LSigFig))
                                    End If
                                End If
                            Else

                                var3 = CDec(NZ(r(i).Item(strCol), 0))
                                var1 = CDec(RoundToDecimalA(var3, LSigFig))

                            End If

                        End If

                        SumSq = SumSq + (var1 - avg) ^ 2
                    End If


                Next i
                k = sk
                If k = 0 Then
                    k = 2
                End If
            End If
        End If
        StdDevDR = Math.Sqrt(SumSq / (k - 1))


    End Function


    Function StdDev(ByVal k, ByVal arr())

        '20170810 LEE:
        'var1 arr(i) comes as double
        'this is resulting in floating point errors

        Dim i As Short
        Dim avg As Object, SumSq As Object
        Dim var1, var2
        Dim d1 As Decimal

        'this function assumes arr(n) is already set at the desired sigfigs
        If k = 1 Then 'can't do sd on n=1
            StdDev = 0
            Exit Function
        End If

        avg = Mean(k, arr)
        For i = 1 To k
            var1 = arr(i) 'debug
            If IsNumeric(var1) Then
                d1 = arr(i)
                SumSq = SumSq + (d1 - avg) ^ 2
            End If

        Next i

        StdDev = Math.Sqrt(SumSq / (k - 1))

    End Function

    Function MeanDR(ByVal r() As DataRow, ByVal strCol As String, ByVal boolAliq As Boolean, ByVal strAliq As String, ByVal boolSF As Boolean, ByVal boolUseIS As Boolean)
        'Finds the mean of a set of data in datarows

        'r - Set of DataRows to collect mean from
        'strCol - Name of Column with set of values to be averaged
        'boolAliq - Whether to divide by the Aliquot factor (i.e. the dilution factor)
        'strAliq - Name of the Column with set of Aliquot factor values
        'boolSF - Whether to use significant figures or not in the Sum (where Mean = Sum / Count)
        'boolUseIS - Not used currently

        '20160510 LEE: Must take in to account NULL values for concentration. Add NZ to conc r(i).Item(strCol)

        Dim Sum As Object
        Dim i As Short
        Dim k As Short
        Dim var1, var2, var3
        Dim varIS, varAnal, varT
        Dim boolHasAnalyte As Boolean = False
        Dim sk As Short

        MeanDR = 0

        Try
            k = r.Length
            Sum = 0
            MeanDR = 0
            If k = 0 Then
                Exit Function
            End If
            If boolSF Then 'prepare data with appropriate sigfigs
                If boolAliq Then
                    '''''''''''console.writeline("Start")
                    For i = 0 To k - 1
                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                        var2 = CDec(r(i).Item(strAliq))

                        var3 = RoundToDecimalA(var1 / var2, LSigFig)
                        var1 = SigFigOrDec(CDec(var3), LSigFig, False) 'debug
                        '''''''''''console.writeline(var1) 'debug
                        Sum = Sum + SigFigOrDec(CDec(var3), LSigFig, False)
                    Next
                    '''''''''''console.writeline("End")
                Else
                    For i = 0 To k - 1
                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                        var3 = RoundToDecimalA(var1, LSigFig)
                        Sum = Sum + SigFigOrDec(CDec(var3), LSigFig, False)
                    Next
                End If
            Else
                If boolAliq Then
                    For i = 0 To k - 1
                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                        var2 = CDec(r(i).Item(strAliq))
                        var3 = var1 / var2
                        Sum = Sum + CDec(var3)
                    Next
                Else
                    sk = 0
                    For i = 0 To k - 1
                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                        If IsNumeric(var1) Then
                            var1 = CDec(var1)
                            sk = sk + 1
                            boolHasAnalyte = True
                            Try
                                varAnal = CDec(NZ(r(i).Item("ANALYTEAREA"), 0))
                            Catch ex As Exception
                                varAnal = 0
                                boolHasAnalyte = False
                            End Try
                            Try
                                varIS = CDec(NZ(r(i).Item("INTERNALSTANDARDAREA"), 0))
                            Catch ex As Exception
                                varIS = 0
                            End Try

                            If StrComp(strCol, "INTERNALSTANDARDAREA", CompareMethod.Text) = 0 Then
                                If boolLUseSigFigsArea Then
                                    var1 = SigFigArea(RoundToDecimalA(varIS, LSigFigArea), LSigFigArea, True, False) 'special rounding incorporated
                                Else
                                    var2 = Format(RoundToDecimalRAFZ(varIS, LSigFigArea), GetRegrDecStr(LSigFigArea))
                                End If
                            Else

                                If boolHasAnalyte Then
                                    If boolRCPARatio Then
                                        If varIS = 0 Then
                                            var1 = 0
                                        Else

                                            'varAnal = SigFigArea(varAnal, LSigFigArea, True, False)
                                            'varIS = SigFigArea(varIS, LSigFigArea, True, False)

                                            If boolLUseSigFigsAreaRatio Then
                                                var1 = SigFigAreaRatio(RoundToDecimalA(varAnal / varIS, LSigFigAreaRatio), LSigFigAreaRatio, True, False) 'special rounding incorporated
                                            Else
                                                var1 = Format(RoundToDecimalRAFZ(varAnal / varIS, LSigFigAreaRatio), GetRegrDecStr(LSigFigAreaRatio))
                                            End If
                                        End If
                                        'ElseIf boolRCPARatio Then
                                        '    If varIS = 0 Then
                                        '        var1 = 0
                                        '    Else
                                        '        var1 = varAnal / varIS ' RoundToDecimalRAFZ(varAnal / varIS, 5)
                                        '    End If
                                    ElseIf boolRCPA Then
                                        'var1 = RoundToDecimalRAFZ(varAnal, 0)
                                        If boolLUseSigFigsArea Then
                                            var1 = SigFigArea(RoundToDecimalA(varAnal, LSigFigArea), LSigFigArea, False, True) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(varAnal, LSigFigArea), GetRegrDecStr(LSigFigArea))
                                        End If
                                        'ElseIf boolRCPA And boolUseIS Then
                                        '    'var1 = RoundToDecimalRAFZ(varIS, 0)
                                    Else
                                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                                        var2 = CDec(NZ(r(i).Item(strAliq), 1))
                                        var3 = var1 / var2 ' SigFigOrDec(RoundToDecimalA(var1 / var2, LSigFig), LSigFig, False)
                                        If boolLUseSigFigs Then
                                            var1 = SigFigOrDec(RoundToDecimalA(var3, LSigFig), LSigFig, False) 'special rounding incorporated
                                        Else
                                            var1 = Format(RoundToDecimalRAFZ(var3, LSigFig), GetRegrDecStr(LSigFig))
                                        End If
                                    End If
                                Else
                                    var3 = NZ(r(i).Item(strCol), 0)
                                    var1 = RoundToDecimalA(var3, LSigFig)
                                End If

                            End If

                            Sum = Sum + CDec(var1)

                        End If

                    Next

                    k = sk
                    If k = 0 Then
                        k = 1
                    End If

                End If
            End If
            MeanDR = Sum / k

        Catch ex As Exception
            'MsgBox("There was a problem calculating a mean for a datarowset.", MsgBoxStyle.Critical, "Invalid data...")
            var1 = ex.Message
            var1 = var1
        End Try

    End Function

    Function MeanDRArea(ByVal r() As DataRow, ByVal strCol As String, ByVal boolAliq As Boolean, ByVal strAliq As String, ByVal boolSF As Boolean, ByVal boolUseIS As Boolean)

        '20160119 LEE:  Deprecated. This function not needed
        '20180722 LEE:  Aack! Still used in some code. Removed reference to MeanDRArea in code. Use MeanDR instead.

        Dim Sum As Object
        Dim i As Short
        Dim k As Short
        Dim var1, var2, var3
        Dim varIS, varAnal, varT

        '20160510 LEE: Must take in to account NULL values for concentration. Add NZ to conc r(i).Item(strCol)

        MeanDRArea = 0

        Try
            k = r.Length
            Sum = 0
            MeanDRArea = 0
            If k = 0 Then
                Exit Function
            End If
            If boolSF Then 'prepare data with appropriate sigfigs
                If boolAliq Then
                    '''''''''''console.writeline("Start")
                    For i = 0 To k - 1
                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                        If boolLUseSigFigsArea Then
                            If LboolWyethRoundingArea Then
                                var1 = RoundToDecimalRAFZ(var1, 0)
                            Else
                                var1 = SigFigArea(var1, LSigFigArea, True, False)
                            End If
                        Else
                            var1 = Format(RoundToDecimalRAFZ(var1, 10), strAreaDec)
                        End If
                        var2 = r(i).Item(strAliq)
                        var3 = CDec(CDec(var1) / CDec(var2)) ' RoundToDecimal(var1 / var2, LSigFigArea + 2)
                        If boolLUseSigFigsArea Then
                            If LboolWyethRoundingArea Then
                                var3 = RoundToDecimalRAFZ(var3, 0)
                            Else
                                var3 = SigFigArea(var3, LSigFigArea, True, False)
                            End If
                        Else
                            var3 = RoundToDecimalRAFZ(var3, LDecArea)
                        End If
                        var1 = CDec(var3) ' SigFigOrDec(CDec(var3), LSigFigArea, True, False) 'debug
                        '''''''''''console.writeline(var1) 'debug
                        'Sum = Sum + SigFigOrDec(CDec(var3), LSigFigArea, True, False)
                        Sum = Sum + var1
                    Next
                    '''''''''''console.writeline("End")
                Else
                    For i = 0 To k - 1
                        var3 = CDec(NZ(r(i).Item(strCol), 0))
                        If boolLUseSigFigsArea Then
                            If LboolWyethRoundingArea Then
                                var3 = RoundToDecimalRAFZ(var3, 0)
                            Else
                                var3 = SigFigArea(var3, LSigFigArea, True, False)
                            End If
                        Else
                            var3 = RoundToDecimalRAFZ(var3, LDecArea)
                        End If
                        'var3 = RoundToDecimal(var1, LSigFigArea + 2)
                        'Sum = Sum + SigFigOrDec(CDec(var3), LSigFigArea, True, False)
                        Sum = Sum + CDec(var3) ' SigFigOrDec(CDec(var3), LSigFigArea, True, False)
                    Next
                End If
            Else
                If boolAliq Then
                    For i = 0 To k - 1
                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                        If boolLUseSigFigsArea Then
                            If LboolWyethRoundingArea Then
                                var1 = RoundToDecimalRAFZ(var1, 0)
                            Else
                                var1 = SigFigArea(var1, LSigFigArea, True, False)
                            End If
                        Else
                            var1 = RoundToDecimalRAFZ(var1, LDecArea)
                        End If
                        var2 = CDec(r(i).Item(strAliq))
                        var3 = var1 / var2
                        Sum = Sum + CDec(var3)
                    Next
                Else
                    For i = 0 To k - 1
                        var1 = CDec(NZ(r(i).Item(strCol), 0))
                        Try
                            varAnal = CDec(NZ(r(i).Item("ANALYTEAREA"), 0))
                        Catch ex As Exception
                            varAnal = 0
                        End Try
                        '20180719 LEE: remove
                        'If boolLUseSigFigsArea Then
                        '    If LboolWyethRoundingArea Then
                        '        'varAnal = RoundToDecimalRAFZ(varAnal, 0)
                        '    Else
                        '        'varAnal = SigFigArea(varAnal, LSigFigArea, True, False)
                        '    End If
                        'Else
                        '    'varAnal = RoundToDecimalRAFZ(varAnal, LDecArea)
                        'End If

                        Try
                            varIS = CDec(NZ(r(i).Item("INTERNALSTANDARDAREA"), 0))
                        Catch ex As Exception
                            varIS = 0
                        End Try
                        '20180719 LEE: remove
                        'If boolLUseSigFigsArea Then
                        '    If LboolWyethRoundingArea Then
                        '        '20180719 LEE: remove
                        '        'varIS = RoundToDecimalRAFZ(varIS, 0)
                        '    Else
                        '        '20180719 LEE: remove
                        '        'varIS = SigFigArea(varIS, LSigFigArea, True, False)
                        '    End If
                        'Else
                        '    'varIS = RoundToDecimalRAFZ(varIS, LDecArea)
                        'End If

                        'If boolRCPARatio And StrComp(strCol, "INTERNALSTANDARDAREA", CompareMethod.Text) = 0 Then
                        'var1 = CDec(varIS) ' RoundToDecimalRAFZ(varIS, 0)
                        'If boolLUseSigFigsAreaRatio Then
                        '    If LboolWyethRoundingArea Then
                        '        var1 = RoundToDecimalRAFZ(varIS, 0)
                        '    Else
                        '        var1 = SigFigAreaRatio(varIS, LSigFigAreaRatio, True, False)
                        '    End If
                        'Else
                        '    var1 = RoundToDecimalRAFZ(varIS, LDecAreaRatio)
                        'End If
                        If boolRCPARatio Then

                            '20180719 LEE:
                            If StrComp(strCol, "INTERNALSTANDARDAREA", CompareMethod.Text) = 0 Then
                                var1 = CDec(varIS)
                                '20180719 LEE:
                                'Here, use LSigFigArea, not LSigFigAreaRatio
                                If boolLUseSigFigsArea Then
                                    If LboolWyethRoundingArea Then
                                        var1 = RoundToDecimalRAFZ(var1, 0)
                                    Else
                                        var1 = SigFigArea(var1, LSigFigArea, True, False)
                                    End If
                                Else
                                    var1 = RoundToDecimalRAFZ(var1, LDecArea)
                                End If
                            Else
                                If varIS = 0 Then
                                    var1 = 0
                                Else
                                    var1 = CDec(CDec(varAnal) / CDec(varIS)) ' RoundToDecimalRAFZ(varAnal / varIS, 5)
                                    If boolLUseSigFigsAreaRatio Then
                                        If LboolWyethRoundingArea Then
                                            var1 = RoundToDecimalRAFZ(var1, 0)
                                        Else
                                            var1 = SigFigAreaRatio(var1, LSigFigAreaRatio, True, False)
                                        End If
                                    Else
                                        var1 = RoundToDecimalRAFZ(var1, LDecAreaRatio)
                                    End If
                                End If
                            End If

                        ElseIf boolRCPA Then
                            If boolUseIS Then
                                var1 = varIS ' RoundToDecimalRAFZ(varIS, 0)
                            Else
                                var1 = varAnal ' RoundToDecimalRAFZ(varAnal, 0)
                            End If
                            'var1 = varAnal ' RoundToDecimalRAFZ(varAnal, 0)
                            If boolLUseSigFigsArea Then
                                If LboolWyethRoundingArea Then
                                    var1 = RoundToDecimalRAFZ(var1, 0)
                                Else
                                    var1 = SigFigArea(var1, LSigFigArea, True, False)
                                End If
                            Else
                                var1 = RoundToDecimalRAFZ(var1, LDecArea)
                            End If
                        Else
                            If boolLUseSigFigsArea Then
                                If LboolWyethRoundingArea Then
                                    var1 = RoundToDecimalRAFZ(var1, 0)
                                Else
                                    var1 = SigFigArea(var1, LSigFigArea, True, False)
                                End If
                            Else
                                var1 = RoundToDecimalRAFZ(var1, LDecArea)
                            End If
                        End If

                        Sum = Sum + CDec(var1)
                    Next
                End If
            End If
            var1 = Sum / k
            'If boolLUseSigFigsArea Then
            '    If LboolWyethRoundingArea Then
            '        var1 = RoundToDecimalRAFZ(var1, 0)
            '    Else
            '        var1 = SigFigArea(var1, LSigFigArea, True, False)
            '    End If
            'Else
            '    var1 = RoundToDecimalRAFZ(var1, LDecArea)
            'End If
            MeanDRArea = var1
        Catch ex As Exception
            'MsgBox("There was a problem calculating a mean for a datarowset.", MsgBoxStyle.Critical, "Invalid data...")

        End Try

    End Function

    Function MeanDRMF(ByVal r() As DataRow, ByVal strCol As String, ByVal boolAliq As Boolean, ByVal strAliq As String, ByVal boolSF As Boolean, ByVal boolUseIS As Boolean)

        'special for matrix effect table

        'Finds the mean of a set of data in datarows

        'r - Set of DataRows to collect mean from
        'strCol - Name of Column with set of values to be averaged
        'boolAliq - Whether to divide by the Aliquot factor (i.e. the dilution factor)  20181202 LEE: DEPRECATED
        'strAliq - Name of the Column with set of Aliquot factor values  20181202 LEE: DEPRECATED
        'boolSF - Whether to use significant figures or not in the Sum (where Mean = Sum / Count)  20181202 LEE: DEPRECATED
        'boolUseIS - Not used currently 20181202 LEE: DEPRECATED

        '20160510 LEE: Must take in to account NULL values for concentration. Add NZ to conc r(i).Item(strCol)

        Dim Sum As Object
        Dim i As Short
        Dim k As Short
        Dim var1, var2, var3
        Dim varIS, varAnal, varT
        Dim boolHasAnalyte As Boolean = False
        Dim sk As Short

        MeanDRMF = 0

        Try
            k = r.Length
            Sum = 0
            MeanDRMF = 0
            If k = 0 Then
                Exit Function
            End If

            sk = 0
            For i = 0 To k - 1
                var1 = CDec(NZ(r(i).Item(strCol), ""))
                If IsNumeric(var1) Then
                    var1 = CDec(var1)
                    sk = sk + 1
                    var1 = NZ(r(i).Item(strCol), 0)
                    Sum = Sum + CDec(var1)

                End If
            Next

            k = sk
            If k = 0 Then
                k = 1
            End If

            MeanDRMF = Sum / k

        Catch ex As Exception
            'MsgBox("There was a problem calculating a mean for a datarowset.", MsgBoxStyle.Critical, "Invalid data...")
            var1 = ex.Message
            var1 = var1
        End Try

    End Function

    Function MeanDV(ByVal dv As System.Data.DataView, ByVal strCol As String, ByVal boolAliq As Boolean, ByVal strAliq As String)
        Dim Sum As Object
        Dim i As Short
        Dim k As Short
        Dim var1, var2, var3

        k = dv.Count
        Sum = 0
        MeanDV = 0
        If k = 0 Then
            Exit Function
        End If
        For i = 0 To k - 1
            'var1 = SigFigOrDec(CDec(dv(i).Item(strCol)), LSigFig, True)
            'var2 = SigFigOrDec(var1, LSigFig, True)
            Sum = Sum + SigFigOrDec(RoundToDecimalA(CDec(dv(i).Item(strCol)), LSigFig), LSigFig, False)
        Next

        If boolAliq Then
            For i = 0 To k - 1
                var1 = dv(i).Item(strCol)
                var2 = dv(1).Item(strAliq)
                var3 = RoundToDecimalA(var1 / var2, LSigFig)
                Sum = Sum + SigFigOrDec(CDec(var3), LSigFig, False)
            Next
        Else
            For i = 0 To k - 1
                var1 = dv(i).Item(strCol)
                var3 = RoundToDecimalA(var1, LSigFig)
                Sum = Sum + SigFigOrDec(CDec(var3), LSigFig, False)
            Next
        End If


        MeanDV = Sum / k

    End Function

    Function GetWt(ByVal intW As Short) As String

        GetWt = "1"

        Select Case intW
            Case 0
                GetWt = "1"
            Case 1
                GetWt = "1/x"
            Case 2
                GetWt = "1/x^2"
            Case 3
                GetWt = "1/y"
            Case 4
                GetWt = "1/y^2"
        End Select

        'Select Case intW
        '    Case 0
        '        GetWt = "1"
        '    Case 1
        '        GetWt = "1/X"
        '    Case 2
        '        GetWt = "1/X^2"
        '    Case 3
        '        GetWt = "1/Y"
        '    Case 4
        '        GetWt = "1/Y^2"
        'End Select


    End Function

    Function GetRegrRegCon(ByVal numRID As Int16, ByVal intAnalyteID As Int32) As String

        GetRegrRegCon = "Linear"

        Dim strF As String
        Dim drows() As DataRow
        Dim dtbl As System.Data.DataTable
        Dim var1

        dtbl = tblRegConAll.Copy

        strF = "RUNID = " & numRID & " AND ANALYTEID = " & intAnalyteID
        drows = dtbl.Select(strF)
        var1 = drows(0).Item("REGRESSIONTEXT")
        GetRegrRegCon = NZ(var1, "Linear")

    End Function

    Function GetWtRegCon(ByVal numRID As Int16, ByVal intAnalyteID As Int32) As String

        GetWtRegCon = "1"

        Dim strF As String
        Dim drows() As DataRow
        Dim dtbl As System.Data.DataTable
        Dim var1
        Dim int1 As Short

        'dtbl = tblRegConAll.Copy
        dtbl = tblRegCon.Copy
        strF = "RUNID = " & numRID & " AND ANALYTEID = " & intAnalyteID
        drows = dtbl.Select(strF)

        If drows.Length = 0 Then
            GetWtRegCon = "NA"
        Else
            Try
                var1 = drows(0).Item("WEIGHTINGFACTOR")
                int1 = NZ(var1, 1)
                GetWtRegCon = GetWt(int1)

            Catch ex As Exception

            End Try
        End If





    End Function

    Function Mean(ByVal k, ByVal arr()) As Decimal

        Mean = 0

        Dim Sum As Object
        Dim i As Integer
        'this function assumes arr(n) is already set at the desired sigfigs
        'arr(0) should be null
        Dim var1
        Dim intK As Short = 0
        Dim d1 As Decimal

        Sum = 0
        Mean = 0
        If k = 0 Then
            Exit Function
        End If
        For i = 1 To k
            var1 = arr(i) 'debug
            If IsNumeric(var1) Then
                d1 = arr(i)
                intK = intK + 1
                Sum = Sum + d1
            End If

        Next i

        If intK = 0 Then
        Else
            Mean = Sum / intK
        End If


    End Function

    Function SigFigOrDecString(ByVal x As Object, ByVal SigFigs As Short, ByVal boolIgnore As Boolean) As String
        If (NZ(x, 0) = 0) Then
            SigFigOrDecString = CStr(DisplayNum(0, SigFigs, boolIgnore))
        Else
            SigFigOrDecString = CStr(DisplayNum(SigFigOrDec(x, SigFigs, boolIgnore), SigFigs, boolIgnore))
        End If
    End Function

    Function SigFigOrDecPeakAreaOrRatio(ByVal x As Object, ByVal SigFigs As Short, ByVal AreaRatio As Boolean)

        '20181203 LEE:

        If AreaRatio Then
            If boolLUseSigFigsAreaRatio Then
                SigFigOrDecPeakAreaOrRatio = SigFigOrDecString(RoundToDecimalA(x, LSigFigAreaRatio), LSigFigAreaRatio, False)
            Else
                SigFigOrDecPeakAreaOrRatio = RoundToDecimalRAFZ(x, LSigFigAreaRatio)
            End If
        Else
            If boolLUseSigFigsArea Then
                SigFigOrDecPeakAreaOrRatio = SigFigOrDecString(RoundToDecimalA(x, LSigFigArea), LSigFigArea, False)
            Else
                SigFigOrDecPeakAreaOrRatio = RoundToDecimalRAFZ(x, LSigFigArea)
            End If
        End If

    End Function

    Function SigFigOrDec(ByVal x As Object, ByVal SigFigs As Short, ByVal boolIgnore As Boolean) As Double

        'Rounds X to Sigfigs significant figures
        'Many thanks to John N. of Locum Destination Consulting for sharing his SigFigOrDec() function for rounding to significant figures.
        'http://www.pcqna.com/Excel_Rounding.htm
        '20150708 Larry: Note the above link doesn't exist anymore
        'Some current (20150708) links (google search for 'SymArith( sig fig')
        'https://ostermiller.org/calc/significant_figures.html
        'http://www.chem.sc.edu/faculty/Morgan/resources/sigfigs/sigfigs6.html:'   this site describes specifically that sigfigs should 'round 5 to even'
        '=> This sigfig function needs to 'round 5 to even'


        'boolIgnore = True means ignore setting of boolLSigFigs

        Dim var1, var2
        Dim num1 As Object
        Dim num2 As Object

        If SigFigs < 1 Then
            SigFigOrDec = 0
            Exit Function
        End If
        '20180722 LEE:
        'WTF???? This evaluation is wrong
        If SigFigs < 0 Then
            SigFigOrDec = x
            Exit Function
        End If

        x = NZ(x, 0)

        If LboolWyethRounding Then 'use wyeth special rounding
            If x >= 99.95 Then
                SigFigOrDec = CDbl(RoundToDecimalRAFZ(x, 0))
                Exit Function
            End If
        End If

        If boolLUseSigFigs Or boolIgnore Then

            Dim Powers As Object, Sign As Int64
            Dim intL As Integer
            Dim boolMkStr As Boolean
            Dim boolNeg As Boolean
            Dim X1 As Object

            If x = 0 Then
                SigFigOrDecString(x, SigFigs, boolIgnore) 'Handle it as string
            Else
                X1 = x
                On Error GoTo ErrHandler

                Sign = Math.Sign(x)
                X1 = Math.Abs(x)
                'sigfig = Sign * vba.RoundToDecimal(X1 / Powers, SigFigs) * Powers
                Powers = 10 ^ (Int(Math.Log10(X1)) + 1) ' * 10 ^ 14
                num1 = X1 / Powers

                'num2 = SymArith(num1, 10 ^ SigFigs)

                '20150629 Larry: this is the culprit. 
                'This function uses a format function that is rounding 34.5 to 35
                'after testing, the Format function used in RoundToDecimal rounds up
                '20150708 Larry: more specifically, 'arithmetic rounding' or 'rounding away from 0'
                num2 = RoundToDecimal(num1, SigFigs)

                SigFigOrDec = CDbl(num2 * Powers * Sign)

            End If
        Else
            SigFigOrDec = RoundToDecimal(x, LDec)
        End If

        Exit Function

ErrHandler:
        SigFigOrDec = 0 'CVErr(xlErrValue)
        'MsgBox(Err.Number & ":  " & Err.Description)
        var1 = var2 'for testing

    End Function

    Function SigFigArea(ByVal x As Object, ByVal SigFigs As Short, ByVal ReturnNumeric As Boolean, ByVal boolIgnore As Boolean) As Object

        'Rounds X to Sigfigs significant figures
        'Many thanks to John N. of Locum Destination Consulting for sharing his SigFigOrDec() function for rounding to significant figures.
        'http://www.pcqna.com/Excel_Rounding.htm
        '20150708 Larry: Note the above link doesn't exist anymore
        'Some current (20150708) links (google search for 'SymArith( sig fig')
        'https://ostermiller.org/calc/significant_figures.html
        'http://www.chem.sc.edu/faculty/Morgan/resources/sigfigs/sigfigs6.html:'   this site describes specifically that sigfigs should 'round 5 to even'
        '=> This sigfig function needs to 'round 5 to even'


        'boolIgnore = True means ignore setting of boolLSigFigs

        Dim var1, var2
        Dim num1 As Object
        Dim num2 As Object

        'If CDbl(x) = 0 Then
        '    SigFig = 0
        '    Exit Function
        'End If


        'If SigFigs < 1 Then
        '    SigFigArea = 0
        '    Exit Function
        'End If
        '20180722 LEE:
        'WTF???? This evaluation is wrong
        If SigFigs < 0 Then
            SigFigArea = x
            Exit Function
        End If

        x = NZ(x, 0)

        If LboolWyethRoundingArea Then 'use wyeth special rounding
            If x >= 99.95 Then
                SigFigArea = RoundToDecimalRAFZ(CDbl(x), 0)
                Exit Function
            End If
        End If

        If boolLUseSigFigsArea Or boolIgnore Then
            Dim Powers As Object, Sign As Int64
            Dim intL As Integer
            Dim boolMkStr As Boolean
            Dim boolNeg As Boolean
            Dim X1 As Object

            If x = 0 Then
                SigFigArea = CDec(RoundToDecimal(CDbl(x), SigFigs - 1))
                ReturnNumeric = False
            Else

                X1 = x
                On Error GoTo ErrHandler

                Sign = Math.Sign(CDbl(x))
                X1 = Math.Abs(CDbl(x))
                'sigfig = Sign * vba.RoundToDecimal(X1 / Powers, SigFigs) * Powers
                Powers = 10 ^ (Int(Math.Log10(X1)) + 1) ' * 10 ^ 14
                num1 = X1 / Powers
                'num2 = SymArith(num1, 10 ^ SigFigs)

                '20150629 Larry: this is the culprit. 
                'This function uses a format function that is rounding 34.5 to 35
                'after testing, the Format function used in RoundToDecimal rounds up
                '20150708 Larry: more specifically, 'arithmetic rounding' or 'rounding away from 0'
                num2 = RoundToDecimal(num1, SigFigs)

                SigFigArea = CDec(num2 * Powers * Sign)

            End If

            If ReturnNumeric Then
            Else
                'SigFig = CStr(DisplaySigFig(SigFig, SigFigs))
                'SigFigArea = CStr(DisplayNum(SigFigArea, SigFigs, boolIgnore))
                SigFigArea = CStr(DisplayNum(SigFigArea, SigFigs, False))
            End If


        Else

            SigFigArea = RoundToDecimal(CDbl(x), LDecArea)
            If ReturnNumeric Then
            Else
                'SigFig = CStr(DisplaySigFig(SigFig, SigFigs))
                'SigFigArea = CStr(DisplayNum(SigFigArea, LDecArea, boolIgnore))
                SigFigArea = CStr(DisplayNum(SigFigArea, LDecArea, True)) 'TRUE gives 'ignore usesigfig boolean
            End If

        End If

        Exit Function

ErrHandler:
        SigFigArea = 0 'CVErr(xlErrValue)
        'MsgBox(Err.Number & ":  " & Err.Description)
        var1 = var2 'for testing

    End Function

    Function SigFigAreaRatio(ByVal x As Object, ByVal SigFigs As Short, ByVal ReturnNumeric As Boolean, ByVal boolIgnore As Boolean) As Object

        'Rounds X to Sigfigs significant figures
        'Many thanks to John N. of Locum Destination Consulting for sharing his SigFigOrDec() function for rounding to significant figures.
        'http://www.pcqna.com/Excel_Rounding.htm
        '20150708 Larry: Note the above link doesn't exist anymore
        'Some current (20150708) links (google search for 'SymArith( sig fig')
        'https://ostermiller.org/calc/significant_figures.html
        'http://www.chem.sc.edu/faculty/Morgan/resources/sigfigs/sigfigs6.html:'   this site describes specifically that sigfigs should 'round 5 to even'
        '=> This sigfig function needs to 'round 5 to even'


        'boolIgnore = True means ignore setting of boolLSigFigs

        Dim var1, var2
        Dim num1 As Object
        Dim num2 As Object

        'If CDbl(x) = 0 Then
        '    SigFig = 0
        '    Exit Function
        'End If

        If SigFigs < 1 Then
            SigFigAreaRatio = 0
            Exit Function
        End If
        '20180722 LEE:
        'WTF???? This evaluation is wrong
        If SigFigs < 0 Then
            SigFigAreaRatio = x
            Exit Function
        End If

        x = NZ(x, 0)

        If LboolWyethRoundingArea Then 'use wyeth special rounding
            If x >= 99.95 Then
                SigFigAreaRatio = RoundToDecimalRAFZ(CDbl(x), 0)
                Exit Function
            End If
        End If

        If boolLUseSigFigsAreaRatio Or boolIgnore Then
            Dim Powers As Object, Sign As Int64
            Dim intL As Integer
            Dim boolMkStr As Boolean
            Dim boolNeg As Boolean
            Dim X1 As Object

            If x = 0 Then
                SigFigAreaRatio = CDec(RoundToDecimal(CDbl(x), SigFigs - 1))
                ReturnNumeric = False
            Else

                X1 = x
                On Error GoTo ErrHandler

                Sign = Math.Sign(CDbl(x))
                X1 = Math.Abs(CDbl(x))
                'sigfig = Sign * vba.RoundToDecimal(X1 / Powers, SigFigs) * Powers
                Powers = 10 ^ (Int(Math.Log10(X1)) + 1) ' * 10 ^ 14
                num1 = X1 / Powers
                'num2 = SymArith(num1, 10 ^ SigFigs)

                '20150629 Larry: this is the culprit. 
                'This function uses a format function that is rounding 34.5 to 35
                'after testing, the Format function used in RoundToDecimal rounds up
                '20150708 Larry: more specifically, 'arithmetic rounding' or 'rounding away from 0'
                num2 = RoundToDecimal(num1, SigFigs)

                SigFigAreaRatio = CDec(num2 * Powers * Sign)

            End If

            If ReturnNumeric Then
            Else
                'SigFig = CStr(DisplaySigFig(SigFig, SigFigs))
                'SigFigArea = CStr(DisplayNum(SigFigArea, SigFigs, boolIgnore))
                SigFigAreaRatio = CStr(DisplayNum(SigFigAreaRatio, SigFigs, False))
            End If


        Else

            SigFigAreaRatio = RoundToDecimal(CDbl(x), LDecAreaRatio)
            If ReturnNumeric Then
            Else
                'SigFig = CStr(DisplaySigFig(SigFig, SigFigs))
                'SigFigAreaRatio = CStr(DisplayNum(SigFigAreaRatio, LDecAreaRatio, boolIgnore))
                SigFigAreaRatio = CStr(DisplayNum(SigFigAreaRatio, LDecAreaRatio, True)) 'TRUE gives 'ignore usesigfig boolean

            End If

        End If

        Exit Function

ErrHandler:
        SigFigAreaRatio = 0 'CVErr(xlErrValue)
        'MsgBox(Err.Number & ":  " & Err.Description)
        var1 = var2 'for testing

    End Function


    Function GetGridColIndex(ByVal strSearch As String, ByVal ts As DataGridTableStyle) As Integer
        Dim c As DataGridColumnStyle
        Dim str1 As String
        Dim Count1 As Integer

        GetGridColIndex = -1
        Count1 = -1
        For Each c In ts.GridColumnStyles
            str1 = c.HeaderText
            Count1 = Count1 + 1
            If StrComp(str1, strSearch, CompareMethod.Text) = 0 Then
                GetGridColIndex = Count1
                Exit For
            End If
        Next
    End Function

    Function AnalRefHook() As String
        Dim strF As String
        Dim rows() As DataRow
        Dim tbl As System.Data.DataTable
        Dim intError As Short

        AnalRefHook = ""
        tbl = tblHooks
        strF = "CHARHOOK = 'CRLWor_AnalRefStandard'"
        rows = tbl.Select(strF)

        If rows.Length = 0 Then 'ignore everything
        Else
            intError = NZ(rows(0).Item("BOOLERROR"), 0)
            If intError = 0 Then 'continue
                AnalRefHook = rows(0).Item("CHARHOOK")
            End If
        End If

    End Function

    Function VerboseNumber(ByVal num As Int16, ByVal boolCap As Boolean) As String

        'bool: True-capitalize first letter, False-nocapitalize first letter

        Dim var1, var2, var3
        Dim ct1 As Short
        Dim Count1 As Short
        Dim arr1(10)
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim num1 As String
        Dim int1 As Short

        VerboseNumber = "zero"
        ct1 = Len(NZ(CStr(num), 0))

        If ct1 = 0 Then
            Exit Function
        End If

        ReDim arr1(ct1)

        If num = 0 Then
            VerboseNumber = "zero"
            GoTo end1
        Else
            For Count1 = 1 To ct1
                arr1(Count1) = Mid(CStr(num), Count1, 1)
            Next
        End If

        If num < 20 And num > 9 Then

            VerboseNumber = GetTeen(num)

        Else
            VerboseNumber = ""
            For Count1 = 1 To ct1
                var1 = arr1(Count1)
                num1 = CInt(Mid(CStr(num), Count1, ct1 - Count1 + 1))

                If num1 < 20 And num1 > 9 Then
                    VerboseNumber = VerboseNumber & " " & GetTeen(num1)
                    Exit For
                Else
                    Select Case var1
                        Case 1
                            str1 = "one"
                        Case 2
                            If num1 < 100 And num1 > 9 Then
                                str1 = "twenty"
                            Else
                                str1 = "two"
                            End If
                        Case 3
                            If num1 < 100 And num1 > 9 Then
                                str1 = "thirty"
                            Else
                                str1 = "three"
                            End If
                        Case 4
                            If num1 < 100 And num1 > 9 Then
                                str1 = "forty"
                            Else
                                str1 = "four"
                            End If
                        Case 5
                            If num1 < 100 And num1 > 9 Then
                                str1 = "fifty"
                            Else
                                str1 = "five"
                            End If
                        Case 6
                            If num1 < 100 And num1 > 9 Then
                                str1 = "sixty"
                            Else
                                str1 = "six"
                            End If
                        Case 7
                            If num1 < 100 And num1 > 9 Then
                                str1 = "seventy"
                            Else
                                str1 = "seven"
                            End If
                        Case 8
                            If num1 < 100 And num1 > 9 Then
                                str1 = "eighty"
                            Else
                                str1 = "eight"
                            End If
                        Case 9
                            If num1 < 100 And num1 > 9 Then
                                str1 = "ninety"
                            Else
                                str1 = "nine"
                            End If
                        Case 0
                            str1 = ""
                    End Select

                    var1 = CInt(Mid(CStr(num), 1, ct1 - Count1 + 1))
                    str2 = ""
                    str3 = ""
                    If Len(str1) = 0 Then

                    Else
                        If var1 < 10 Then
                            str2 = ""
                            str3 = str1
                        ElseIf var1 < 100 Then
                            str2 = ""
                            str3 = str1
                        ElseIf var1 < 1000 Then
                            str2 = "hundred"
                            str3 = str1 & " " & str2
                        ElseIf var1 < 10000 Then
                            str2 = "thousand"
                            str3 = str1 & " " & str2
                        End If
                        If ct1 > 1 Then
                            If Count1 = ct1 Then
                                VerboseNumber = Trim(VerboseNumber & "-" & str3)
                            Else
                                VerboseNumber = Trim(VerboseNumber & " " & str3)
                            End If
                        Else
                            VerboseNumber = Trim(VerboseNumber & " " & str3)
                        End If
                    End If
                End If
            Next
            VerboseNumber = Trim(VerboseNumber)

        End If

end1:

        If boolCap Then 'capitalize
            str1 = Mid(VerboseNumber, 1, 1)
            int1 = Asc(str1)
            str2 = Chr(int1 - 32)
            str3 = str2 & Mid(VerboseNumber, 2, Len(VerboseNumber) - 1)
            VerboseNumber = str3
        End If

    End Function

    Function GetTeen(ByVal num As Int16) As String

        Select Case num
            Case 10
                GetTeen = "ten"
            Case 11
                GetTeen = "eleven"
            Case 12
                GetTeen = "twelve"
            Case 13
                GetTeen = "thirteen"
            Case 14
                GetTeen = "fourteen"
            Case 15
                GetTeen = "fifteen"
            Case 16
                GetTeen = "sixteen"
            Case 17
                GetTeen = "seventeen"
            Case 18
                GetTeen = "eighteen"
            Case 19
                GetTeen = "nineteen"
        End Select


    End Function


    Function WorkingPageWidth(ByVal wd As Microsoft.Office.Interop.Word.Application) As Single
        Dim lm, rm, pw, w
        With wd.Selection.PageSetup
            lm = .LeftMargin ' = InchesToPoints(1.4)
            rm = .RightMargin ' = InchesToPoints(1)
            pw = .PageWidth ' = InchesToPoints(8.6)
        End With

        w = pw - lm - rm
        WorkingPageWidth = w

    End Function

    Function GetPDFDriver() As String

        Dim tbl As System.Data.DataTable
        Dim var1
        Dim rows() As DataRow
        Dim strF As String

        tbl = tblConfiguration
        strF = "CHARCONFIGTITLE = 'PDF Print Driver'"
        rows = tbl.Select(strF)
        var1 = NZ(rows(0).Item("CHARCONFIGVALUE"), "")

        GetPDFDriver = var1

    End Function

    Function ReturnDirectoryBrowse(ByVal boolFileName As Boolean, ByVal path As String, ByVal strFilter As String, ByVal strFileName As String, boolNewFolder As Boolean) As String

        Dim var1
        Dim str1 As String
        Dim int1 As Short
        Dim intLen As Short
        Dim Count1 As Short

        ReturnDirectoryBrowse = ""

        If boolFileName Then 'use file bowser

            frmH.OpenFileDialog1.Multiselect = False
            frmH.OpenFileDialog1.InitialDirectory = path
            frmH.OpenFileDialog1.Filter = strFilter '"All files (*.*)|*.*"
            frmH.OpenFileDialog1.FilterIndex = 1
            frmH.OpenFileDialog1.FileName = strFileName '"*.*"
            frmH.OpenFileDialog1.Title = "Browse to file..."
            If frmH.OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                ReturnDirectoryBrowse = frmH.OpenFileDialog1.FileName
            End If
        Else 'use folder browser
            frmH.FolderBrowserDialog1.SelectedPath = path
            frmH.FolderBrowserDialog1.ShowNewFolderButton = boolNewFolder
            If frmH.FolderBrowserDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                ReturnDirectoryBrowse = frmH.FolderBrowserDialog1.SelectedPath
                'ensure last character is "\"
                str1 = Mid(ReturnDirectoryBrowse, Len(ReturnDirectoryBrowse), 1)
                If StrComp(str1, "\", CompareMethod.Text) = 0 Then
                Else
                    ReturnDirectoryBrowse = ReturnDirectoryBrowse & "\"
                End If
            End If

        End If


    End Function

    '*********** Code Start **********
    'This code was originally written by Dev Ashish
    'It is not to be altered or distributed,
    'except as part of an application.
    'You are free to use it in any application,
    'provided the copyright notice is left unchanged.
    '
    'Code Courtesy of
    'Dev Ashish
    '
    Function fWhoHasDocFileOpen(ByVal strDocFile As String) As String
        '*******************************************
        'Name:      fWhoHasDocFileOpen (Function)
        'Purpose:   Returns the network name of the user who has
        '              strDocFile open
        'Author:     Dev Ashish
        'Date:        February 11, 1999, 07:28:13 PM
        'Called by: Any
        'Calls:       fFileDirPath
        'Inputs:      strDocFile - Complete path to the Word document
        'Output:     Name of the user if successful,
        '               vbNullString on error
        '*******************************************
        On Error GoTo ErrHandler
        Dim intFree As Integer
        Dim intPos As Integer
        Dim strDoc As String
        Dim strFile As String
        Dim strExt As String
        Dim strUserName As String

        intFree = FreeFile()
        strDoc = Dir(strDocFile)
        intPos = InStr(1, strDoc, ".")
        If intPos > 0 Then
            strFile = Left$(strDoc, intPos - 1)
            strExt = Right$(strDoc, Len(strDoc) - intPos)
        End If
        intPos = 0
        If Len(strFile) > 6 Then
            If Len(strFile) = 7 Then
                strDocFile = fFileDirPath(strDocFile) & "~$" & _
                  Mid$(strFile, 2, Len(strFile)) & "." & strExt
            Else
                strDocFile = fFileDirPath(strDocFile) & "~$" & _
                  Mid$(strFile, 3, Len(strFile)) & "." & strExt
            End If
        Else
            strDocFile = fFileDirPath(strDocFile) & "~$" & Dir(strDocFile)
        End If


        Dim fn1 As Short
        Dim fn2 As Short
        Dim str1 As String

        fn1 = FreeFile()
        FileOpen(fn1, strDocFile, OpenMode.Input, , OpenShare.Shared)
        str1 = LineInput(1)
        strUserName = Right$(str1, Len(str1) - 1)
        fWhoHasDocFileOpen = strUserName

        'Open strDocFile For Input Shared As #intFree
        'Line Input #intFree, strUserName
        '      strUserName = Right$(strUserName, Len(strUserName) - 1)
        '      fWhoHasDocFileOpen = strUserName
ExitHere:
        On Error Resume Next
        'Close #intFree
        FileClose(fn1)
        If Err.Number <> 0 Then
            Dim var1
            var1 = "oo"
        End If

        Exit Function
ErrHandler:
        fWhoHasDocFileOpen = vbNullString
        Resume ExitHere
    End Function

    Function fFileDirPath(ByVal strFile As String) As String
        'Code courtesy of
        'Terry Kreft & Ken Getz
        Dim strPath As String
        strPath = Dir(strFile)
        fFileDirPath = Left(strFile, Len(strFile) - Len(strPath))
    End Function

    Function GetLegendTitle(ByVal idCRT As Int16, ByVal idRT As Int16) As String

        GetLegendTitle = ""

        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim rows() As DataRow
        Dim strFirst As String

        Try
            tbl = tblTableProperties
            strF = "ID_TBLREPORTTABLE = " & idRT
            rows = tbl.Select(strF)
            GetLegendTitle = NZ(rows(0).Item("CHARTITLELEG"), "")
            If Len(GetLegendTitle) = 0 Then
                If BOOLDIFFERENCE Then
                    GetLegendTitle = "%Difference"
                ElseIf boolRECOVERY Then
                    GetLegendTitle = "%Recovery"
                ElseIf boolMEANACCURACY Then
                    GetLegendTitle = "Mean Accuracy"
                End If
            End If
        Catch ex As Exception

        End Try


    End Function


    Function EvalCoverPage(wd As Microsoft.Office.Interop.Word.Application) As Boolean

        EvalCoverPage = False

        With wd

            Dim int1 As Int32
            Dim str1 As String

            Try
                str1 = .ActiveDocument.Characters(1).Text
                int1 = AscW(str1)
            Catch ex As Exception

            End Try

            '20170618 LEE: Do not set styles to 'Normal' anymore. Next character may have a different style that is there for a reason
            If int1 = 13 Then

                str1 = .ActiveDocument.Characters(2).Text
                int1 = AscW(str1)

                If int1 = 12 Then
                    .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
                    '.Selection.Style = .ActiveDocument.Styles("Normal")
                    EvalCoverPage = True
                End If

            ElseIf int1 = 12 Then

                .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then
                Else
                    '.Selection.Style = .ActiveDocument.Styles("Normal")
                End If

                EvalCoverPage = True

            End If

        End With

    End Function

    Function BQLVerbose() As String

        BQLVerbose = "Below Quantitation Limit"

        Dim str1 As String

        str1 = NZ(gstrBQL, "BQL/AQL")

        Select Case str1
            Case "BQL/AQL", "AQL/BQL"
                BQLVerbose = "Below Quantitation Limit"
            Case "BLQ/ALQ", "ALQ/BLQ"
                BQLVerbose = "Below Limit of Quantitation"
            Case "LLOQ/ULOQ"
                BQLVerbose = "Lower Limit of Quantitation"
        End Select

    End Function

    Function AQLVerbose() As String

        AQLVerbose = "Above Quantitation Limit"

        Dim str1 As String

        str1 = NZ(gstrBQL, "BQL/AQL")

        Select Case str1
            Case "BQL/AQL", "AQL/BQL"
                AQLVerbose = "Above Quantitation Limit"
            Case "BLQ/ALQ", "ALQ/BLQ"
                AQLVerbose = "Above Limit of Quantitation"
            Case "LLOQ/ULOQ"
                AQLVerbose = "Upper Limit of Quantitation"
        End Select

    End Function

    Function BQL() As String

        BQL = "BQL"

        Dim str1 As String

        str1 = NZ(gstrBQL, "BQL/AQL")

        Select Case str1
            Case "BQL/AQL", "AQL/BQL"
                BQL = "BQL"
            Case "BLQ/ALQ", "ALQ/BLQ"
                BQL = "BLQ"
            Case "LLOQ/ULOQ"
                BQL = "LLOQ"
        End Select

    End Function

    Function AQL() As String

        AQL = "AQL"

        Dim str1 As String

        str1 = NZ(gstrBQL, "BQL/AQL")

        Select Case str1
            Case "BQL/AQL", "AQL/BQL"
                AQL = "AQL"
            Case "BLQ/ALQ", "ALQ/BLQ"
                AQL = "ALQ"
            Case "LLOQ/ULOQ"
                AQL = "ULOQ"
        End Select

    End Function

    Function GetNewCBS() As Int64

        GetNewCBS = 0

        'find intcbs
        Dim strF As String = "CHARTITLE = '" & strReportTemplateChoice & "'"
        Dim rowsCBS() As DataRow = tblWordStatements.Select(strF)
        If rowsCBS.Length = 0 Then
            GetNewCBS = frmH.dgvReportStatements("ID_TBLWORDSTATEMENTS", 0).Value
        Else
            GetNewCBS = rowsCBS(0).Item("ID_TBLWORDSTATEMENTS")
        End If

        'Dim str1 As String
        'Dim dvRR As System.Data.DataView = frmH.dgvReportStatementWord.DataSource
        'Dim CountRR As Short
        'For CountRR = 0 To dvRR.Count - 1
        '    str1 = dvRR(CountRR).Item("CHARTITLE")
        '    If StrComp(str1, strReportTemplateChoice, CompareMethod.Text) = 0 Then
        '        GetNewCBS = dvRR(CountRR).Item("ID_TBLWORDSTATEMENTS")
        '        Exit For
        '    End If
        'Next

    End Function

    Function NBHReal() As String

        'there are some instances when it is OK to use a real nonbreaking hyphen, like in report body

        'Wiki:  http://en.wikipedia.org/wiki/Hyphen
        'NBH: Soft hyphen. Optional. IF IS A MINUS SIGN, IT WILL NOT PRINT!!!
        '       http://www.fileformat.info/info/unicode/char/ad/index.htm
        'chrw(2011): Hard hyphen. Non-breaking-hyphen
        '       http://www.fileformat.info/info/unicode/char/2011/index.htm
        'chrw(8209): 

        'normal hypen = chrw(45)

        '20140226 Gubbs: This is causing too many problems. Merck Intervet Shawn called today and minus signs are not getting printed.
        ' will do only in Analyte Name and others described in Sub ReturnSearch
        '20150805 Larry: Nope, have to get rid of it entirely

        '20151221 LEE: Here's the scoop with NBH
        'When StudyDoc creates document with NBH (chrw(173)), the original document produced has this problem:
        'If user copy/pastes into another document, clipboard converts 173 to 31, which is an invisible hyphen (e.g. not visible when printed).
        'To resolve:
        '    First save, then close the StudyDoc document, then re-open
        '    Clipboard will now preserve 173 as 173
        'But still don't do this in tables, it's too risky
        'Only do it when search/replace field codes (e.g. accuracymin..., accuracymax..., etc) that mainly occur in the report body which most likely won't get copy/pasted right awway.
        'So it's done in:
        '   - Sub ReturnAnova
        '   - sub ExecuteA

        NBHReal = ChrW(173)

    End Function

    Function fNBSP() As String

        fNBSP = ChrW(160)

    End Function

    Function NBH() As String

        'don't do anymore
        'Exit Sub

        'Wiki:  http://en.wikipedia.org/wiki/Hyphen
        'NBH: Soft hyphen. Optional. IF IS A MINUS SIGN, IT WILL NOT PRINT!!!
        '       http://www.fileformat.info/info/unicode/char/ad/index.htm
        'chrw(2011): Hard hyphen. Non-breaking-hyphen
        '       http://www.fileformat.info/info/unicode/char/2011/index.htm
        'chrw(8209): 

        'normal hypen = chrw(45)

        '20140226 Gubbs: This is causing too many problems. Merck Intervet Shawn called today and minus signs are not getting printed.
        ' will do only in Analyte Name and others described in Sub ReturnSearch
        '20150805 Larry: Nope, have to get rid of it entirely

        '20151221 LEE: Here's the scoop with NBH
        'When StudyDoc creates document with NBH (chrw(173)), the original document produced has this problem:
        'If user copy/pastes into another document, clipboard converts 173 to 31, which is an invisible hyphen (e.g. not visible when printed).
        'To resolve:
        '    First save, then close the StudyDoc document, then re-open
        '    Clipboard will now preserve 173 as 173
        'But still don't do this in tables, it's too risky
        'Only do it when search/replace field codes (e.g. accuracymin..., accuracymax..., etc) that mainly occur in the report body which most likely won't get copy/pasted right awway.
        'So it's done in:
        '   - Sub ReturnAnova
        '   - sub ExecuteA

        NBH = ChrW(45) ' ChrW(173)

    End Function

    Function Allowed(strCol As String) As Boolean

        '20160711 LEE:
        'This function deprecated. All global permissions values are now set at login: Sub SetPermissions


        Allowed = False

        Dim strF As String
        Dim rows() As DataRow
        Dim int1 As Short

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        rows = tblPermissions.Select(strF)

        'NDL: Note - if no user selected (e.g. Cancel button pressed on login), there will be no rows in the table
        Allowed = False
        If (rows.Length > 0) Then
            int1 = NZ(rows(0).Item(strCol), 0) 'make 0 default
            If (int1 <> 0) Then
                Allowed = True
            End If
        End If
end1:

    End Function

    Function IsGuWuFast(wd As Microsoft.Office.Interop.Word.Application) As Boolean

        IsGuWuFast = False

        With wd

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            .Selection.Find.ClearFormatting()
            With .Selection.Find
                .Text = "DoGuWuFast"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue ' wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute()

                If Not .Found Then
                    IsGuWuFast = False
                Else
                    IsGuWuFast = True
                End If

            End With

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        End With

    End Function

    Function SaveAsDocx(wd As Microsoft.Office.Interop.Word.Application) As Boolean

        SaveAsDocx = False

        With wd

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            .Selection.Find.ClearFormatting()
            With .Selection.Find
                .Text = "StudyDoc save as docx"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue ' wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute()

                If Not .Found Then
                    SaveAsDocx = False
                Else
                    SaveAsDocx = True
                End If

            End With

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        End With

    End Function

    Function RandomPswd() As String

        Dim Count1 As Integer
        Dim var1, var2
        Dim RV
        Dim UB As Short = 122
        Dim LB As Short = 49

        Dim str1 As String = ""

        Dim int1 As Integer = 0

        Do Until int1 = 16

            Randomize() 'call randomize to seed Rnd with a different number between application logins
            var1 = CInt(Int((UB - LB + 1) * Rnd() + LB))

            If (var1 >= 58 And var1 <= 64) Or (var1 >= 91 And var1 <= 96) Then
            Else
                int1 = int1 + 1
                var2 = ChrW(var1)
                str1 = str1 & var2
            End If

        Loop

        RandomPswd = str1

    End Function

    Function ReturnOutlierMethod() As String

        ReturnOutlierMethod = "(Outlier Method not specified)"

        Dim dtbl As DataTable = tblData
        Dim strF As String
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim rows() As DataRow = dtbl.Select(strF)
        Dim str1 As String
        str1 = NZ(rows(0).Item("CHAROUTLIERMETHOD"), "(No Outlier Method specified)")
        ReturnOutlierMethod = str1

    End Function


    Function CheckColLenEx(ByVal str1 As String, ByVal intCL As Short, strMod As String, strSource As String) As Boolean

        CheckColLenEx = False

        Dim strM As String
        Dim intA As Short

        Try
            If Len(str1) > intCL Then

                CheckColLenEx = True

                strM = "This field is limited in length to " & intCL & "." & ChrW(10) & ChrW(10)
                strM = strM & "The length of the entered text is " & Len(str1) & "." & ChrW(10) & ChrW(10)
                strM = strM & "Please modify the text to conform to the defined text limit."
                strM = strM & ChrW(10) & ChrW(10) & strMod & " - " & strSource & " cell"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            Else
                CheckColLenEx = False
            End If
        Catch ex As Exception

        End Try


    End Function

    Function GetColumnName(ByVal strColDescr As String, ByVal strtbl As String) As String

        GetColumnName = ""

        Dim str1 As String

        Dim dtbl As System.Data.DataTable
        Dim strF As String
        Dim rows() As DataRow

        dtbl = tblDataTableRowTitles

        strF = "CHARROWNAME = '" & strColDescr & "' AND CHARTABLEREF = '" & strtbl & "'"
        rows = dtbl.Select(strF)

        If rows.Length = 0 Then
            GetColumnName = ""
        Else
            GetColumnName = NZ(rows(0).Item("CHARTABLEREFCOLUMNNAME"), "")
        End If

    End Function

    Function GetCL(ByVal strTbl As String, ByVal strColName As String) As Short

        'legend:
        'Public boolGuWuAccess As Boolean = False
        'Public boolGuWuSQLServer As Boolean = False
        'Public boolGuWuOracle As Boolean = False

        GetCL = 255

        If boolGuWuAccess Or boolGuWuSQLServer Then

            GetCL = 255

            Select Case strTbl
                Case "tblMethodValidationData", UCase("tblMethodValidationData")

                    Select Case strColName
                        Case "CHARCORPORATESTUDYID"
                            GetCL = 255
                        Case "CHARPROTOCOLNUMBER"
                            GetCL = 255
                        Case "CHARMETHODVALIDATIONTITLE"
                            GetCL = 255
                        Case "CHARVALREPORTNUM"
                            GetCL = 200
                        Case "CHARLMTITLE"
                            GetCL = 255
                        Case "CHARLMNUMBER"
                            GetCL = 255
                        Case "CHARANALMETHODTYPE"
                            GetCL = 200
                        Case "CHARSPONSORMETHODVALIDATIONID"
                            GetCL = 255
                        Case "CHARSPONSORMETHVALTITLE"
                            GetCL = 255
                        Case "CHARASSAYDESCRIPTION"
                            GetCL = 255
                        Case "CHARSAMPLESIZEUNITS"
                            GetCL = 255
                        Case "CHARSPECIES"
                            GetCL = 255
                        Case "CHARANTICOAGULANT"
                            GetCL = 255
                        Case "CHARMATRIX"
                            GetCL = 255
                        Case "CHARMAXRUNSIZE"
                            GetCL = 255
                        Case "CHARQCCONC"
                            GetCL = 250
                        Case "CHARCALIBRCONC"
                            GetCL = 250
                        Case "CHARLLOQ"
                            GetCL = 25
                        Case "CHARULOQ"
                            GetCL = 25
                        Case "CHARAVERECANAL"
                            GetCL = 25
                        Case "CHARAVERECIS"
                            GetCL = 25
                        Case "CHARINTERQCACCRNG"
                            GetCL = 25
                        Case "CHARINTERQCPRECRNG"
                            GetCL = 50
                        Case "CHARINTRAQCACCRNG"
                            GetCL = 50
                        Case "CHARINTRAQCPRECRNG"
                            GetCL = 50
                        Case "CHARDEMONSTRATEDFREEZETHAW"
                            GetCL = 255
                        Case "CHARMAXNUMBERFREEZETHAW"
                            GetCL = 255
                        Case "CHARSTABILITYUNDERSTORAGECOND"
                            GetCL = 255
                        Case "CHARSTABILITYMAXSTORAGEDUR"
                            GetCL = 255
                        Case "CHARPROCSTABILITY"
                            GetCL = 255
                        Case "CHARREFRSTAB"
                            GetCL = 255
                        Case "CHARLTSTORSTAB"
                            GetCL = 255
                        Case "CHARDILINTEGR"
                            GetCL = 255
                        Case "CHARANALSELECT"
                            GetCL = 250
                        Case "CHARISSELECT"
                            GetCL = 250
                        Case "CHARFTSTORCOND"
                            GetCL = 200

                        Case "CHARBLOOD"
                            GetCL = 250
                        Case "CHARSTOCKSOLUTION"
                            GetCL = 250
                        Case "CHARSPIKING"
                            GetCL = 250
                        Case "CHARAUTOSAMPLER"
                            GetCL = 250
                        Case "CHARBATCHREINJECTION"
                            GetCL = 250

                        Case Else
                            GetCL = -1
                    End Select

                Case "tblData", UCase("tblData")

                    Select Case strColName
                        Case "charCorporateStudyID", UCase("charCorporateStudyID")
                            GetCL = 255
                        Case "charProtocolNumber", UCase("charProtocolNumber")
                            GetCL = 255
                        Case "charSponsorStudyNumber", UCase("charSponsorStudyNumber")
                            GetCL = 255
                        Case "charSponsorStudyTitle", UCase("charSponsorStudyTitle")
                            GetCL = 255
                        Case "DTSTUDYSTARTDATE"
                            GetCL = -1
                        Case "DTSTUDYENDDATE"
                            GetCL = -1
                        Case "charDataArchivalLocation", UCase("charDataArchivalLocation")
                            GetCL = 255
                        Case "NUMSIGFIGS"
                            GetCL = -1
                        Case "NUMDECIMALS"
                            GetCL = -1
                        Case "BOOLUSESIGFIGS"
                            GetCL = -1
                        Case "BOOLUSESPECRND"
                            GetCL = -1
                        Case "NUMREGRSIGFIGS"
                            GetCL = -1
                        Case Else
                            GetCL = -1
                    End Select

                Case "tblCompanyAnalRefTable", UCase("tblCompanyAnalRefTable")

                    Select Case strColName

                        Case "CHARCOMPANYID"
                            GetCL = 255
                        Case "CHARANALYTENAME"
                            GetCL = 255
                        Case "CHARIUPAC"
                            GetCL = 255
                        Case "CHARALIAS"
                            GetCL = 255
                        Case "CHARCHEMSTRUCTURE"
                            GetCL = 100
                        Case "CHARMOLFORMULA"
                            GetCL = 100
                        Case "CHARMOLWT"
                            GetCL = 100
                        Case "CHARMONOISOWT"
                            GetCL = 100
                        Case "CHARLOTNUMBER"
                            GetCL = 255
                        Case "CHARPHYSICALDESCRIPTION"
                            GetCL = 255
                        Case "CHARSTORAGECONDITIONS"
                            GetCL = 255
                        Case "DTDATERECEIVED"
                            GetCL = 255
                        Case "DTEXPIRATIONRETESTDATE"
                            GetCL = 255
                        Case "CHARAMOUNTRECEIVED"
                            GetCL = 255
                        Case "CHARSUPPLIER"
                            GetCL = 255
                        Case "CHARPURITY"
                            GetCL = 255
                        Case "CHARPERCENTWATER"
                            GetCL = 255
                        Case "BOOLISCOADMINISTERED"
                            GetCL = -1
                        Case "CHARCERTOFANALYSIS"
                            GetCL = 255
                        Case "CHARCOMMENTS"
                            GetCL = 255
                        Case "BOOLISREPLICATE"
                            GetCL = -1
                        Case "BOOLWATSON"
                            GetCL = -1
                        Case "ID_TBLANALREFSTANDARDS"
                            GetCL = -1
                        Case "CHARANALYTEPARENT"
                            GetCL = 255
                        Case "BOOLISINTSTD"
                            GetCL = -1

                    End Select

                Case "tblFieldCodes", UCase("tblFieldCodes")
                    Select Case strColName
                        Case "CHARFIELDCODE"
                            GetCL = 255
                        Case "CHARDESCRIPTION"
                            GetCL = 255
                        Case "CHAREXAMPLE"
                            GetCL = 255
                    End Select

                Case "tblCustomFieldCodes", UCase("tblCustomFieldCodes")
                    Select Case strColName
                        Case "CHARVALUE"
                            GetCL = 255
                    End Select

                Case "tblFieldCodes", UCase("tblFieldCodes")
                    Select Case strColName
                        Case "CHARFIELDCODE"
                            GetCL = 255
                        Case "CHARFIELDCODE"
                            GetCL = 2000
                        Case "CHARFIELDCODE"
                            GetCL = 2000
                    End Select

            End Select

        Else
            GetCL = 2000
        End If

    End Function

    Function boolCLExceeded(ByVal strColText As String, ByVal strTbl As String, ByVal strStr As String, ByVal boolColName As Boolean, strMod As String, strSource As String) As Boolean

        boolCLExceeded = False
        'MethValColumnLimit

        Dim strColName As String
        Dim boolCL As Boolean = False

        Try
            strColName = ""
            If boolColName Then
                strColName = strColText
            Else
                strColName = GetColumnName(strColText, strTbl)
            End If

            If Len(strColName) = 0 Then
                boolCLExceeded = False
                GoTo end1
            End If

            Dim intCL As Short

            intCL = -1

            intCL = GetCL(strTbl, strColName)


            Dim intStrL As Short
            Dim strM As String

            If intCL = -1 Then
                boolCLExceeded = False
            Else
                intStrL = Len(strStr)
                If intStrL > intCL Then
                    boolCLExceeded = True

                    strM = "This field is limited in length to " & intCL & "." & ChrW(10) & ChrW(10)
                    strM = strM & "The length of the entered text is " & intStrL & "." & ChrW(10) & ChrW(10)
                    strM = strM & "" & strStr & "" & ChrW(10) & ChrW(10)
                    strM = strM & "Please modify the text to conform to the defined text limit."
                    strM = strM & ChrW(10) & ChrW(10) & strMod & " - " & strSource & " cell"
                    MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")

                Else
                    boolCLExceeded = False
                End If
            End If
        Catch ex As Exception

        End Try

end1:

    End Function


    Function GetLastNumber(ByVal strType As String) As Short

        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intRows As Short
        Dim strF As String

        Select Case strType

            Case "Tables"
                dtbl = tblReportTableAnalytes
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE <> 0"
            Case "Figures"
                dtbl = tblAppFigs
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDEINREPORT <> 0 AND BOOLFIGURE <> 0"
            Case "Appendices", "Attachments" 'NOTE: Attachments depricated
                dtbl = tblAppFigs
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDEINREPORT <> 0 AND BOOLAPPENDIX <> 0"

        End Select

        rows = dtbl.Select(strF)
        intRows = rows.Length

        GetLastNumber = intRows.ToString

    End Function


    Function GetNameID(ByVal idUID As Int64)

        'need to find User Account and UserID
        'get user id from tblPermissions
        Dim idUA As Int64
        'Dim idUA As Int64
        Dim strUID As String
        Dim strUA As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strF As String
        Dim rowsT() As DataRow


        'idUA = dr.Item("ID_TBLUSERACCOUNTS")
        Erase rowsT
        strF = "ID_TBLUSERACCOUNTS = " & idUID
        rowsT = tblUserAccounts.Select(strF)
        strUID = rowsT(0).Item("CHARUSERID")
        idUA = rowsT(0).Item("ID_TBLPERSONNEL")
        Erase rowsT
        strF = "ID_TBLPERSONNEL = " & idUA
        rowsT = tblPersonnel.Select(strF)
        str1 = NZ(rowsT(0).Item("CHARFIRSTNAME"), "")
        str2 = NZ(rowsT(0).Item("CHARMIDDLENAME"), "")
        str3 = NZ(rowsT(0).Item("CHARLASTNAME"), "")

        If Len(str2) = 0 Then
            strUA = str1 & " " & str3
        Else
            strUA = str1 & " " & str2 & " " & str3
        End If

        Dim arr1(2) As String
        arr1(1) = strUA
        arr1(2) = strUID

        GetNameID = arr1


    End Function

    Function GetSpecial(ByVal strItem As String, ByVal strTName As String, ByVal varVal As Object, ByVal strTblDescr As String) As String

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim var1

        GetSpecial = CStr(varVal)

        Select Case strTName

            Case "TBLUSERACCOUNTS"
                Select Case strItem
                    Case "CHARPASSWORD"
                        GetSpecial = "*****"
                    Case "ID_TBLWATSONACCOUNT"
                        'get watson account from 
                        strF = "USERID = " & varVal
                        rows = tblWatsonUsers.Select(strF)
                        If rows.Length = 0 Then
                            var1 = varVal & " (No Watson Account found)"
                        Else
                            var1 = varVal & " (" & rows(0).Item("LOGINNAME") & ")"
                        End If
                        GetSpecial = var1
                End Select
            Case "TBLMETHODVALIDATIONDATA"
                Select Case strItem
                    Case "ID_TBLSTUDIES2"
                        'varVal returns ID. Must convert to string
                        If IsDBNull(varVal) Then
                        Else
                            strF = "ID_TBLSTUDIES = " & varVal
                            rows = tblStudies.Select(strF)
                            'GetSpecial = rows(0).Item("CHARWATSONSTUDYNAME") & " (ID_TBLSTUDIES = " & varVal & ")"
                            If rows.Length = 0 Then
                                GetSpecial = "[None]"
                            Else
                                GetSpecial = rows(0).Item("CHARWATSONSTUDYNAME") & " (ID_TBLSTUDIES = " & varVal & ")"
                            End If
                        End If
                End Select

            Case "TBLTEMPLATES"
                Select Case strItem
                    Case "ID_TBLSTUDIES"
                        'varVal returns ID. Must convert to string
                        If IsDBNull(varVal) Then
                        Else
                            strF = "ID_TBLSTUDIES = " & varVal
                            rows = tblStudies.Select(strF)
                            'GetSpecial = rows(0).Item("CHARWATSONSTUDYNAME") & " (ID_TBLSTUDIES = " & varVal & ")"
                            If rows.Length = 0 Then
                                GetSpecial = "Not Applicable"
                            Else
                                GetSpecial = rows(0).Item("CHARWATSONSTUDYNAME") & " (ID_TBLSTUDIES = " & varVal & ")"
                            End If
                        End If
                End Select

            Case "TBLREPORTS"
                Select Case strItem
                    Case "BOOLEXCLUDEPSAE"
                        If varVal = 0 Then
                            GetSpecial = "Report All Analytical Runs"
                        ElseIf varVal = 1 Then
                            GetSpecial = "Exclude PSAE Analytical Runs"
                        ElseIf varVal = 2 Then
                            GetSpecial = "Report All Accepted Analytical Runs"
                        End If

                    Case "BOOLALLAR"
                        GetSpecial = "Report All Analytical Runs"
                    Case "BOOLACCAR"
                        GetSpecial = "Include Accepted Analytical Runs"
                    Case "BOOLREJAR"
                        GetSpecial = "Include Rejected Analytical Runs"
                    Case "BOOLREGRAR"
                        GetSpecial = "Include 'Regression Performed' Analytical Runs"
                    Case "BOOLNOREGRAR"
                        GetSpecial = "Include 'NO Regression Performed' Analytical Runs"
                    Case "BOOLINCLPSAE"
                        GetSpecial = "Include PSAE Analytical Runs"
                    Case "ID_TBLCONFIGREPORTTYPE"
                        'get report type from tblConfigReportType
                        tbl = tblConfigReportType
                        strF = "ID_TBLCONFIGREPORTTYPE = " & varVal
                        rows = tbl.Select(strF)
                        GetSpecial = NZ(rows(0).Item("CHARREPORTTYPE"), "Sample Analysis")
                    Case "INTUSERCOMMENTS"
                        If varVal = 1 Then
                            GetSpecial = "Use Watson Comments"
                        ElseIf varVal = 2 Then
                            GetSpecial = "User User Comments"
                        End If
                    Case "BOOLMULTIVALSUM"
                        If varVal = 0 Then
                            GetSpecial = "Single Table"
                        Else
                            GetSpecial = "Table for Each Analyte"
                        End If
                End Select
            Case "TBLDATA", "tblData"
                Select Case strItem
                    Case "ID_TBLASSAYTECHNIQUE"
                        'get appropriate value
                        tbl = tblDropdownBoxContent
                        strF = "ID_TBLDROPDOWNBOXCONTENT = " & varVal
                        rows = tbl.Select(strF)
                        GetSpecial = rows(0).Item("CHARVALUE")
                    Case "ID_TBLANTICOAGULANT"
                        'get appropriate value
                        tbl = tblDropdownBoxContent
                        strF = "ID_TBLDROPDOWNBOXCONTENT = " & varVal
                        rows = tbl.Select(strF)
                        GetSpecial = rows(0).Item("CHARVALUE")
                    Case "ID_SUBMITTEDBY", "ID_SUBMITTEDTO", "ID_INSUPPORTOF"
                        'get appropriate value
                        tbl = tblCorporateNickNames
                        strF = "ID_TBLCORPORATENICKNAMES = " & varVal
                        rows = tbl.Select(strF)
                        If rows.Length = 0 Then
                            GetSpecial = "NA"
                        Else
                            GetSpecial = rows(0).Item("CHARNICKNAME")
                        End If

                End Select

            Case "TBLREPORTSTATEMENTS"
                Select Case strItem
                    Case "ID_TBLWORDSTATEMENTS"

                        'find id in tblWordStatements
                        tbl = tblWordStatements
                        strF = "ID_TBLWORDSTATEMENTS = " & varVal
                        rows = tbl.Select(strF)
                        If rows.Length = 0 Then
                            GetSpecial = "NA"
                        Else
                            GetSpecial = rows(0).Item("CHARTITLE")
                        End If

                End Select

            Case "TBLCONFIGURATION"
                If InStr(1, strTblDescr, "Authentication Type", CompareMethod.Text) > 0 Then
                    If IsNumeric(varVal) Then
                        var1 = CInt(varVal)
                        Select Case var1
                            Case 1
                                GetSpecial = varVal & "  (Description: LDAP)"
                            Case 2
                                GetSpecial = varVal & "  (Description: Non-LDAP)"
                            Case 3
                                GetSpecial = varVal & "  (Description: ADVAPI32)"
                        End Select
                    End If
                End If


        End Select


    End Function

    Function ExcludeItem(ByVal strItem As String, ByVal strTName As String, ByVal CountT As Short) As Boolean

        ExcludeItem = False

        If CountT = 1 Then 'continue
        Else
            Exit Function
        End If

        'DON'T HAVE TO EXCLUDE FOR ID_TBL. THE FOR/NEXT LOOP STARTS AT 1, EXCLUDING ID_TBL
        'Select Case strTName
        '    Case "TBLTEMPLATES"
        '    Case Else
        '        Dim strID As String
        '        If StrComp(strItem, "ID_TBLSTUDIES2", CompareMethod.Text) = 0 Then
        '        Else
        '            strID = Mid(strItem, 1, 3)
        '            If StrComp(strID, "ID_", CompareMethod.Text) = 0 Then
        '                ExcludeItem = True
        '                GoTo end1
        '            End If
        '        End If
        'End Select


        Select Case strItem
            Case "BOOLA", "BOOLI", "boolA", "boolI", "boolS", "UPSIZE_TS", "BOOLAPP", "BOOLIR", "BOOLFIG"
                ExcludeItem = True
                GoTo end1
        End Select

        Select Case strTName

            Case "TBLMETHODVALIDATIONDATA"

                'determine if sample analysis
                Dim strT As String
                Dim rowsT() As DataRow
                Dim strF As String

                strF = "ID_TBLSTUDIES = " & id_tblStudies
                rowsT = tblReports.Select(strF)
                strT = NZ(rowsT(0).Item("CHARREPORTTYPE"), "Sample Analysis")
                If InStr(1, strT, "Sample", CompareMethod.Text) > 0 Then
                    'allow only id_tblstudies2
                    Select Case strItem
                        Case "ID_TBLSTUDIES2"
                        Case Else
                            ExcludeItem = True
                    End Select
                Else
                    Select Case strItem
                        Case "CHARARCHIVEPATH"
                            ExcludeItem = True
                        Case "CHARCORPORATESTUDYID"
                            ExcludeItem = True
                        Case "CHARPROTOCOLNUMBER"
                            ExcludeItem = True
                        Case "CHARMETHODVALIDATIONTITLE"
                            ExcludeItem = True
                        Case "CHARVALREPORTNUM"
                            ExcludeItem = True
                        Case "CHARANALMETHODTYPE"
                            ExcludeItem = True
                    End Select
                End If

            Case "TBLHOOKS"

                Select Case strItem
                    Case "BOOLERROR"
                        ExcludeItem = True
                End Select

            Case "TBLCONFIGURATION"
                Select Case strItem
                    Case "CHARCONFIGVALUE"
                    Case Else
                        ExcludeItem = True
                End Select
            Case "TBLREASONFORCHANGE"
                Select Case strItem
                    Case "DEFAULTCHK"
                        ExcludeItem = True
                End Select

            Case "TBLMEANINGOFSIG"
                Select Case strItem
                    Case "DEFAULTCHK"
                        ExcludeItem = True
                End Select

            Case "TBLQATABLES"
                Select Case strItem
                    Case "INTORDER"
                        ExcludeItem = True
                    Case "ID_TBLREPORTTABLEHEADERCONFIG"
                        ExcludeItem = True

                End Select
            Case "TBLCONTRIBUTINGPERSONNEL"
                Select Case strItem
                    Case "boolIncludeSOTP"
                        ExcludeItem = True
                End Select

            Case "TBLREPORTSTATEMENTS"
                Select Case strItem
                    Case "CHARSECTIONNAME"
                        ExcludeItem = True
                    Case "charSectionName"
                        ExcludeItem = True
                    Case "boolI"
                        ExcludeItem = True
                    Case "boolPB"
                        ExcludeItem = True
                    Case "CHARSTATEMENT"
                        ExcludeItem = True
                End Select

            Case "TBLSAMPLERECEIPT"
                Select Case strItem
                    Case "boolU"
                        ExcludeItem = True
                    Case "INTORDER"
                        ExcludeItem = True
                End Select

            Case "TBLSUMMARYDATA"
                Select Case strItem
                    Case "boolI"
                        ExcludeItem = True
                    Case "CHARVALUE"
                        ExcludeItem = True
                End Select

            Case "TBLREPORTS"
                Select Case strItem
                    Case "INTCALSTD"
                        ExcludeItem = True
                    Case "INTQC"
                        ExcludeItem = True
                    Case "INTSHOWBQL"
                        ExcludeItem = True
                    Case "INTSHOWCALSTD"
                        ExcludeItem = True
                    Case "ID_TBLCONFIGREPORTTYPE"
                        ExcludeItem = True
                End Select

            Case "TBLRREPORTTABLE"
                Select Case strItem
                    Case "CHARSTYLE"
                        ExcludeItem = True
                End Select

            Case "TBLTABLEPROPERTIES"
                Select Case strItem
                    Case "BOOLSTATSNR"
                        ExcludeItem = True
                    Case "BOOLSTATSLETTER"
                        ExcludeItem = True
                    Case "BOOLCSREPORTACCVALUES"
                        ExcludeItem = True
                End Select

            Case "TBLSAMPLERECEIPT"
                Select Case strItem
                    Case "INTORDER"
                        ExcludeItem = True
                End Select

            Case "TBLPERSONNEL"
                Select Case strItem
                    Case "DTACTIVATED"
                        ExcludeItem = True
                    Case "DTDEACTIVATED"
                        ExcludeItem = True
                    Case "boolA"
                        ExcludeItem = True
                End Select

            Case "TBLUSERACCOUNTS"
                Select Case strItem
                    Case "DTACTIVATED"
                        ExcludeItem = True
                    Case "DTDEACTIVATED"
                        ExcludeItem = True
                    Case "DTTIMESTAMP"
                        ExcludeItem = True
                End Select

            Case "TBLREPORTTABLEHEADERCONFIG"

            Case "TBLCORPORATENICKNAMES"
                Select Case strItem
                    Case "boolI"
                        ExcludeItem = True
                End Select

            Case "TBLREPORTTABLEANALYTES"
                Select Case strItem
                    Case "NUMINCSAMPLECRIT01"
                    Case "BOOLINCLUDE"
                    Case Else
                        ExcludeItem = True
                End Select

        End Select

end1:

    End Function


    Function boolAllowAcc(ByVal idT As Int64) As Boolean

        boolAllowAcc = False

        Dim id As Int64
        Dim strM As String

        id = idT

        If id = 2 Or id = 13 Or id = 14 Or id = 15 Or id = 22 Or id = 23 Or id = 33 Or id = 34 Or id = 35 Then
            boolAllowAcc = False
        Else
            boolAllowAcc = True
        End If

    End Function

    Function FindCStuff(ByVal strVal As String, ByVal strAnal As String) As String

        FindCStuff = "[None]"

        If Len(strVal) = 0 Then
            GoTo end1
        End If

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short

        str1 = " for "
        str2 = " for " & strAnal

        int1 = InStr(1, strVal, str1, CompareMethod.Text)
        int2 = InStr(1, strVal, str2, CompareMethod.Text)

        If int1 = 0 And int2 = 0 Then
            FindCStuff = strVal
            GoTo end1
        End If

        If int1 = int2 Then 'first position
            FindCStuff = Mid(strVal, 1, int1 - 1)
        Else 'somewhere in middle
            int1 = 1
            int4 = 1
            Do Until int1 = int2
                int3 = int1 - 1
                int1 = InStr(int4, strVal, str1, CompareMethod.Text)
                int4 = int1 + 1

            Loop

            FindCStuff = Mid(strVal, int3 + Len(str2) + 2, int2 - (int3 + Len(str2) + 2))

        End If

end1:

    End Function

    Function FindSumRow(ByVal strName As String, ByVal tbl As System.Data.DataTable) As Short

        FindSumRow = -1

        Dim Count1 As Short
        Dim str1 As String

        For Count1 = 0 To tbl.Rows.Count - 1
            str1 = tbl.Rows(Count1).Item("Item")
            If StrComp(str1, strName, CompareMethod.Text) = 0 Then
                FindSumRow = Count1
                Exit Function
            End If
        Next

    End Function

    Function AppendixLetter(ByVal int1)
        'will convert number to letter
        Dim var1, var2

        'Capital A = 65
        var1 = Chr(65 + int1 - 1)
        'var1 = ChrW(var1)
        AppendixLetter = var1

    End Function

    Function BoolAccCrit(ByVal varVal As Object) As Boolean

        BoolAccCrit = False

        If IsDBNull(varVal) Then
            BoolAccCrit = True
            GoTo end1
        End If

        If Len(varVal) = 0 Then
            BoolAccCrit = True
            GoTo end1
        End If

        If IsNumeric(varVal) Then
        Else
            GoTo end1
        End If

        If varVal >= 0 Then
        Else
            GoTo end1
        End If

        BoolAccCrit = True

end1:
        If BoolAccCrit Then
        Else
            Dim str1 As String

            str1 = "Entry must be numeric >= 0" & ChrW(10)
            'str1 = str1 & "If you wish to stop this error message, un-check the 'Use StudyDoc Acceptance Criteria' box"
            MsgBox(str1, MsgBoxStyle.Information, "Invalid entry...")

        End If

    End Function

    Function IsFig(ByVal strE) As Boolean

        IsFig = False

        'strExt = "*.jpg,*.jpeg,*.bmp,*.tif,*.tiff"
        If StrComp(strE, ".jpg", CompareMethod.Text) = 0 Then
            IsFig = True
        ElseIf StrComp(strE, ".jpeg", CompareMethod.Text) = 0 Then
            IsFig = True
        ElseIf StrComp(strE, ".bmp", CompareMethod.Text) = 0 Then
            IsFig = True
        ElseIf StrComp(strE, ".tif", CompareMethod.Text) = 0 Then
            IsFig = True
        ElseIf StrComp(strE, ".tiff", CompareMethod.Text) = 0 Then
            IsFig = True
        ElseIf StrComp(strE, ".png", CompareMethod.Text) = 0 Then
            IsFig = True
        End If

    End Function

    Function GetExt(ByVal strE As String) As String

        GetExt = ".a"

        Dim Count1 As Short
        Dim str1 As String

        For Count1 = Len(strE) To 1 Step -1

            str1 = Mid(strE, Count1, 1)
            If StrComp(str1, ".", CompareMethod.Text) = 0 Then
                Exit For
            End If

        Next

        If Count1 = 1 Then
        Else
            GetExt = Mid(strE, Count1, Len(strE))
        End If

    End Function


    Function SearchReplaceRunIDTest(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strFind As String) As Boolean ', ByVal intS, ByVal intE)

        SearchReplaceRunIDTest = True

        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        'Dim myRng As Microsoft.Office.Interop.Word.selection
        Dim myRng As Microsoft.Office.Interop.Word.Range
        Dim strM As String
        Dim strM1 As String
        Dim strM2 As String
        Dim boolFound As Boolean
        Dim mySel As Microsoft.Office.Interop.Word.Selection

        Try
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        Catch ex As Exception

            GoTo end1

        End Try

        myRng = wd.Selection.Range
        mySel = wd.Selection

        Dim strR As String

        strM = frmH.lblProgress.Text


        'first determine if there is something to replace
        With mySel.Find
            .ClearFormatting()
            '.Text = strFind
            .Forward = True
            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
            .Execute(FindText:=strFind)

            If .Found Then
                SearchReplaceRunIDTest = True
            Else

                '20180726 LEE:
                'Need to check for potential IS
                Dim strFind1 As String = "IS_" & Mid(strFind, 2, Len(strFind))
                With mySel.Find
                    .ClearFormatting()
                    '.Text = strFind
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Execute(FindText:=strFind1)

                    If .Found Then
                        SearchReplaceRunIDTest = True
                    Else
                        SearchReplaceRunIDTest = False
                    End If

                End With

            End If

        End With

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

end1:

    End Function


    Function GetDateFromRunID(ByVal idRun As Int16, ByVal strFormat As String, ByVal intGroup As Short, ByVal idRT As Int32) As String

        GetDateFromRunID = "NA"

        Dim str1 As String
        Dim strF As String
        Dim var1, var2, var3
        Dim int2 As Int32

        '20180823 LEE:
        'need to see if this data comes from external data
        strF = "INTGROUP = " & intGroup & " AND RUNID = " & idRun & " AND ID_TBLREPORTTABLE = " & idRT
        Dim rowsAS() As DataRow = tblAssignedSamples.Select(strF)
        int2 = rowsAS.Length
        If int2 = 0 Then
            GetDateFromRunID = GetDate(idRun, strFormat)
        Else 'evaluate id_tblstudy and id_tblStudy2
            var1 = NZ(rowsAS(0).Item("ID_TBLSTUDIES"), 0)
            var2 = NZ(rowsAS(0).Item("ID_TBLSTUDIES2"), 0)
            If var1 = var2 Then
                GetDateFromRunID = GetDate(idRun, strFormat)
            Else
                var3 = NZ(rowsAS(0).Item("ASSAYDATETIME"), "")
                If Len(var3) = 0 Then
                    GetDateFromRunID = GetDate(idRun, strFormat)
                Else
                    GetDateFromRunID = Format(rowsAS(0).Item("ASSAYDATETIME"), strFormat)
                End If
            End If
        End If


    End Function

    Function GetDate(ByVal idRun As Int16, strFormat As String) As String

        Dim str1 As String
        Dim rowsAR() As DataRow
        Dim strF1 As String
        Dim var1, var2


        strF1 = "[Watson Run ID] = '" & CStr(idRun) & "'"
        rowsAR = tblAnalRunSum.Select(strF1)

        If rowsAR.Length = 0 Then
            GetDate = "NA"
        Else
            var1 = rowsAR(0).Item("Analysis Date")
            var2 = NZ(var1, "NA")
            If StrComp(var2, "NA", CompareMethod.Text) = 0 And IsDate(var2) = False Then
                GetDate = "NA"
            Else
                Try
                    GetDate = Format(CDate(var2), strFormat)
                Catch ex As Exception
                    GetDate = var2
                End Try
            End If
        End If

    End Function

    Function GetQCDec() As Short

        Dim row() As DataRow
        Dim strF As String
        Dim var1

        GetQCDec = gintQCDec

        Try
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            row = tblData.Select(strF)
            var1 = row(0).Item("INTQCPERCDECPLACES")

            GetQCDec = NZ(var1, gintQCDec)

        Catch ex As Exception

        End Try


    End Function

    Function GetQCDecStr() As String

        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        GetQCDecStr = "0.0"

        int1 = intQCDec

        If int1 = 0 Then
            GetQCDecStr = "0"
        Else
            str1 = "0."
            For Count1 = 1 To int1
                str1 = str1 & "0"
            Next
            GetQCDecStr = str1
        End If


    End Function

    Function GetRegrDecStr(int1) As String

        'Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        GetRegrDecStr = "0.0"

        'int1 = LRegrDec

        If int1 = 0 Then
            GetRegrDecStr = "0"
        Else
            str1 = "0."
            For Count1 = 1 To int1
                str1 = str1 & "0"
            Next
            GetRegrDecStr = str1
        End If


    End Function

    Function GetAreaDecStr() As String

        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        GetAreaDecStr = "0.0"

        int1 = LDecArea

        If int1 = 0 Then
            GetAreaDecStr = "0"
        Else
            str1 = "0."
            For Count1 = 1 To int1
                str1 = str1 & "0"
            Next
            GetAreaDecStr = str1
        End If


    End Function

    Function GetAreaRatioDecStr() As String

        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        GetAreaRatioDecStr = "0.0"

        int1 = LDecAreaRatio

        If int1 = 0 Then
            GetAreaRatioDecStr = "0"
        Else
            str1 = "0."
            For Count1 = 1 To int1
                str1 = str1 & "0"
            Next
            GetAreaRatioDecStr = str1
        End If


    End Function


    Function getIntTTot(boolSingleTable As Boolean) As Short

        Dim dgvT As DataGridView = frmH.dgvReportTableConfiguration
        Dim intRowT As Short
        Dim intRowsT As Short
        Dim intRowA As Short
        Dim intRowsA As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim idT As Int64
        Dim idCT As Int64
        Dim boolIS As Boolean
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim strF4 As String
        Dim strF5 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolAddIS As Boolean



        getIntTTot = 0

        Try
            If boolSingleTable Then

                intRowT = dgvT.CurrentRow.Index

                idT = dgvT("ID_TBLREPORTTABLE", intRowT).Value
                idCT = dgvT("ID_TBLCONFIGREPORTTABLES", intRowT).Value
                boolAddIS = False

                'determine if table allows intstd
                strF1 = "ID_TBLCONFIGREPORTTABLES =" & idCT
                Dim rows1() As System.Data.DataRow = tblConfigReportTables.Select(strF1)
                int1 = rows1(0).Item("BOOLINCLUDEIS")
                If int1 <> 0 Then 'continue looking
                    'find out if table is configured to use intstd
                    strF2 = "ID_TBLCONFIGREPORTTABLES =" & idCT & " AND ID_TBLREPORTTABLE = " & idT
                    Dim rows2() As System.Data.DataRow = tblTableProperties.Select(strF2)
                    int2 = rows2(0).Item("BOOLINCLUDEISTBL")
                    If int2 <> 0 Then
                        'add IS to total
                        boolAddIS = True
                    End If
                End If

                'now add tables
                Dim rows3() As System.Data.DataRow ' 
                strF3 = "IsIntStd = 'No'"
                Select Case idCT

                    Case 1, 22, 23
                        '20180717 LEE:
                        'these tables look at parent compound regardless of matrix
                        Dim dv1 As DataView = New DataView(tblAnalytesHome, strF3, "", DataViewRowState.CurrentRows)
                        Dim tblA As DataTable = dv1.ToTable("a", True, "ANALYTEID")
                        Dim dv2 As DataView = dgvT.DataSource
                        Dim tblT As DataTable = dv2.ToTable
                        For Count1 = 0 To tblA.Rows.Count - 1
                            strF4 = "ANALYTEID = " & tblA.Rows(Count1).Item("ANALYTEID")
                            Dim rows4() As System.Data.DataRow = tblAnalytesHome.Select(strF4)
                            For Count2 = 0 To rows4.Length - 1
                                str1 = rows4(Count2).Item("ANALYTEDESCRIPTION")
                                If dgvT.Columns.Contains(str1) Then
                                    int3 = dgvT(str1, intRowT).Value
                                    If int3 <> 0 Then
                                        getIntTTot = getIntTTot + 1
                                        If boolIS Then
                                            getIntTTot = getIntTTot + 1
                                        End If
                                        Exit For
                                    End If
                                End If

                            Next Count2

                        Next Count1



                    Case 5, 6, 7, 30
                        'these tables look at parent compound
                        '20160502 LEE: Need to account for matrix as well
                        Dim dv1 As DataView = New DataView(tblAnalytesHome, strF3, "", DataViewRowState.CurrentRows)
                        Dim tblA As DataTable = dv1.ToTable("a", True, "ANALYTEID", "MATRIX")
                        Dim dv2 As DataView = dgvT.DataSource
                        Dim tblT As DataTable = dv2.ToTable
                        For Count1 = 0 To tblA.Rows.Count - 1
                            strF4 = "ANALYTEID = " & tblA.Rows(Count1).Item("ANALYTEID") & " AND MATRIX ='" & tblA.Rows(Count1).Item("MATRIX") & "'"
                            Dim rows4() As System.Data.DataRow = tblAnalytesHome.Select(strF4)
                            For Count2 = 0 To rows4.Length - 1
                                str1 = rows4(Count2).Item("ANALYTEDESCRIPTION")
                                If dgvT.Columns.Contains(str1) Then
                                    int3 = dgvT(str1, intRowT).Value
                                    If int3 <> 0 Then
                                        getIntTTot = getIntTTot + 1
                                        If boolIS Then
                                            getIntTTot = getIntTTot + 1
                                        End If
                                        Exit For
                                    End If
                                End If

                            Next Count2

                        Next Count1

                    Case Else
                        'rows3 = tblAnalytesHome.Select(strF3)
                        If boolAddIS Then
                            strF3 = "IsIntStd = 'No' or IsIntStd = 'Yes'"
                        Else
                            strF3 = "IsIntStd = 'No'"
                        End If
                        rows3 = tblAnalytesHome.Select(strF3)
                        For Count1 = 0 To rows3.Length - 1
                            str1 = rows3(Count1).Item("ANALYTEDESCRIPTION")
                            str2 = rows3(Count1).Item("IsIntStd")
                            If boolAddIS And StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                                getIntTTot = getIntTTot + 1
                            Else
                                If dgvT.Columns.Contains(str1) Then
                                    int3 = dgvT(str1, intRowT).Value
                                    If int3 <> 0 Then
                                        getIntTTot = getIntTTot + 1
                                        If boolIS Then
                                            getIntTTot = getIntTTot + 1
                                        End If
                                    End If
                                End If
                            End If

                        Next
                End Select

            Else

                For Count3 = 0 To dgvT.RowCount - 1

                    intRowT = Count3

                    idT = dgvT("ID_TBLREPORTTABLE", intRowT).Value
                    idCT = dgvT("ID_TBLCONFIGREPORTTABLES", intRowT).Value
                    boolAddIS = False

                    'determine if table allows intstd
                    strF1 = "ID_TBLCONFIGREPORTTABLES =" & idCT
                    Dim rows1() As System.Data.DataRow = tblConfigReportTables.Select(strF1)
                    int1 = rows1(0).Item("BOOLINCLUDEIS")
                    If int1 <> 0 Then 'continue looking
                        'find out if table is configured to use intstd
                        strF2 = "ID_TBLCONFIGREPORTTABLES =" & idCT & " AND ID_TBLREPORTTABLE = " & idT
                        Dim rows2() As System.Data.DataRow = tblTableProperties.Select(strF2)

                        '20171205 LEE:
                        If rows2.Length = 0 Then
                        Else
                            int2 = rows2(0).Item("BOOLINCLUDEISTBL")
                            If int2 <> 0 Then
                                'add IS to total
                                boolAddIS = True
                            End If
                        End If

                    End If

                    'now add tables
                    Dim rows3() As System.Data.DataRow ' 
                    strF3 = "IsIntStd = 'No'"
                    Select Case idCT
                        Case 5, 6, 7, 22, 23, 30
                            'these tables look at parent compound
                            Dim dv1 As DataView = New DataView(tblAnalytesHome, strF3, "", DataViewRowState.CurrentRows)
                            Dim tblA As DataTable = dv1.ToTable("a", True, "ANALYTEID", "MATRIX")
                            Dim dv2 As DataView = dgvT.DataSource
                            Dim tblT As DataTable = dv2.ToTable
                            For Count1 = 0 To tblA.Rows.Count - 1
                                strF4 = "ANALYTEID = " & tblA.Rows(Count1).Item("ANALYTEID") & " AND MATRIX ='" & tblA.Rows(Count1).Item("MATRIX") & "'"
                                Dim rows4() As System.Data.DataRow = tblAnalytesHome.Select(strF4)
                                For Count2 = 0 To rows4.Length - 1
                                    str1 = rows4(Count2).Item("ANALYTEDESCRIPTION")
                                    If dgvT.Columns.Contains(str1) Then
                                        int3 = dgvT(str1, intRowT).Value
                                        If int3 <> 0 Then
                                            getIntTTot = getIntTTot + 1
                                            If boolIS Then
                                                getIntTTot = getIntTTot + 1
                                            End If
                                            Exit For
                                        End If
                                    End If

                                Next Count2

                            Next Count1

                        Case Else

                            'rows3 = tblAnalytesHome.Select(strF3)
                            'For Count1 = 0 To rows3.Length - 1
                            '    str1 = rows3(Count1).Item("ANALYTEDESCRIPTION")
                            '    If dgvT.Columns.Contains(str1) Then
                            '        int3 = dgvT(str1, intRowT).Value
                            '        If int3 <> 0 Then
                            '            getIntTTot = getIntTTot + 1
                            '            If boolIS Then
                            '                getIntTTot = getIntTTot + 1
                            '            End If
                            '        End If
                            '    End If
                            'Next

                            If boolAddIS Then
                                strF3 = "IsIntStd = 'No' or IsIntStd = 'Yes'"
                            Else
                                strF3 = "IsIntStd = 'No'"
                            End If
                            rows3 = tblAnalytesHome.Select(strF3)
                            For Count1 = 0 To rows3.Length - 1
                                str1 = rows3(Count1).Item("ANALYTEDESCRIPTION")
                                str2 = rows3(Count1).Item("IsIntStd")
                                If boolAddIS And StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                                    getIntTTot = getIntTTot + 1
                                Else
                                    If dgvT.Columns.Contains(str1) Then
                                        int3 = dgvT(str1, intRowT).Value
                                        If int3 <> 0 Then
                                            getIntTTot = getIntTTot + 1
                                            If boolIS Then
                                                getIntTTot = getIntTTot + 1
                                            End If
                                        End If
                                    End If
                                End If

                            Next


                    End Select

                Next Count3

            End If

        Catch ex As Exception

        End Try



    End Function

    Function GetSString() As String

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim intCt As Short

        GetSString = ""

        For Count1 = 1 To intSort

            strF = arrSort(3, Count1)
            strS = arrSort(2, Count1)

            If Len(GetSString) = 0 Then
                GetSString = strF & " " & strS
            Else
                GetSString = GetSString & ", " & strF & " " & strS
            End If

        Next

        'For Count1 = 1 To intGroups

        '    strF = arrSort(3, Count1)
        '    strS = arrSort(2, Count1)

        '    If Len(GetSString) = 0 Then
        '        GetSString = strF & " " & strS
        '    Else
        '        GetSString = GetSString & ", " & strF & " " & strS
        '    End If

        'Next

    End Function

    Function GetGString() As String

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim intCt As Short

        GetGString = ""

        For Count1 = 1 To intGroups

            strF = arrGroups(3, Count1)
            strS = arrGroups(2, Count1)

            If Len(GetGString) = 0 Then
                GetGString = strF & " " & strS
            Else
                GetGString = GetGString & ", " & strF & " " & strS
            End If

        Next

        'For Count1 = 1 To intGroups

        '    strF = arrSort(3, Count1)
        '    strS = arrSort(2, Count1)

        '    If Len(GetGString) = 0 Then
        '        GetGString = strF & " " & strS
        '    Else
        '        GetGString = GetGString & ", " & strF & " " & strS
        '    End If

        'Next

    End Function

    Function GetCaption(ByVal strReport As String) As String

        Dim strF As String
        Dim rowP() As DataRow
        Dim rowU() As DataRow
        Dim rowPerm() As DataRow
        Dim strPerm As String
        Dim tblP As System.Data.DataTable
        Dim tblU As System.Data.DataTable
        Dim tblPerm As System.Data.DataTable
        Dim idPerm As Int64
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String

        tblP = tblPersonnel
        tblU = tblUserAccounts
        tblPerm = tblPermissions

        'find user account
        strF = "id_tblUserAccounts = " & id_tblUserAccounts

        'NDL: If no user yet (e.g. Cancel button pressed on login), then don't show anything
        '20160711 LEE: No, must show normal heading with no user
        'If (IsNothing(rowU)) Then
        If id_tblUserAccounts < 1 Then

            str1 = ""

            'MeCaption = ""
            'Exit Function

        Else

            rowU = tblU.Select(strF)
            str1 = rowU(0).Item("charUserID")

            idPerm = rowU(0).Item("ID_TBLPERMISSIONS")
            rowPerm = tblPerm.Select("ID_TBLPERMISSIONS = " & idPerm)
            strPerm = rowPerm(0).Item("CHARPERMISSIONSNAME")

            'find user
            strF = "id_tblPersonnel = " & id_tblPersonnel
            rowP = tblP.Select(strF)
            If id_tblPersonnel = 0 Then

            Else
                str2 = rowP(0).Item("charFirstName")
                str3 = NZ(rowP(0).Item("charMiddleName"), "")
                str4 = rowP(0).Item("charLastName")
                If Len(str3) = 0 Then
                    str5 = str2 & " " & str4
                Else
                    str5 = str2 & " " & str3 & " " & str4
                End If
            End If


        End If


        Select Case strReport
            Case "ReportWriter"
                'str2 = "Gubbs Inc GuWu" & ChrW(174) & " - Report Writer"
                str2 = "LABIntegrity StudyDoc" & ChrW(8482) & " - Report Writing Manager"
            Case "Console"
                str2 = GetStudyDocHeader(False)
        End Select


        If id_tblUserAccounts < 1 Then
            gUserLabel = " - User: No User Logged In"
            str2 = str2 & " v" & GetVersion() & gUserLabel
        Else

            If gboolLDAP Then
                str3 = NZ(rowU(0).Item("CHARNETWORKACCOUNT"), "NA")
                gUserLabel = " - UserName: " & str5 & " logged in as Network UserID: " & str3
            Else
                gUserLabel = " - UserName: " & str5 & " logged in as StudyDoc UserID: " & str1
            End If
            str2 = str2 & " v" & GetVersion() & gUserLabel & " assigned to Permissions Group: " & strPerm
        End If

        GetCaption = str2
        MeCaption = GetCaption

        gUserName = str5
        gUserID = str1


    End Function

    Function GetVersion() As String

        Dim var1
        Dim int1 As Short
        Dim int2 As Short

        var1 = System.Windows.Forms.Application.ProductVersion
        GetVersion = var1
        int1 = 1
        Do Until InStr(int1, var1, ".", CompareMethod.Text) = 0
            int2 = InStr(int1, var1, ".", CompareMethod.Text)
            If int2 = 0 Then
                Exit Do
            End If
            int1 = int2 + 1
        Loop
        var1 = Mid(var1, 1, int1 - 2)
        GetVersion = var1


    End Function

    Function GetVersionFour() As String

        Dim var1
        Dim int1 As Short
        Dim int2 As Short

        var1 = System.Windows.Forms.Application.ProductVersion
        GetVersionFour = var1


    End Function


    Function CreatePDF(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strPath As String) As String

        CreatePDF = strPath
        'convert path to .pdf
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strP As String
        Dim int1 As Short
        Dim strM As String
        Dim strM1 As String
        Dim var1


        Dim numP As Short


        For Count1 = Len(strPath) To 1 Step -1
            str1 = Mid(strPath, Count1, 1)
            If StrComp(str1, ".", CompareMethod.Text) = 0 Then
                int1 = Count1
                Exit For
            End If
        Next

        str2 = Mid(strPath, int1, Len(strPath) - int1 + 1)
        Dim strP1 As String = strPath
        strP = Replace(strPath, str2, ".PDF", 1, -1, CompareMethod.Text)

        Try
            'wd.ActiveDocument.ExportAsFixedFormat(OutputFileName:=strP, ExportFormat:=Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, _
            '    OpenAfterExport:=False, OptimizeFor:=Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Range:= _
            '    Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, From:=1, To:=1, Item:=Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, _
            '    IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
            '    Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
            '    True, UseISO19005_1:=False)

            Try
                var1 = wd.Version
            Catch ex As Exception
                var1 = var1
            End Try
            If var1 < 12 Then 'must have office 2007 or greater
                strM = "Word 2007 or later must be installed on this workstation" & ChrW(10) & ChrW(10)
                strM1 = strM & "PDF not created..." & ChrW(10) & ChrW(10)

                MsgBox(strM1, MsgBoxStyle.Information, "PDF not created...")
                gDoPDF = False
            Else

                'frmH.lblProgress.Text = "Saving PDF:" & ChrW(10) & ChrW(10) & strP
                'frmH.lblProgress.Refresh()

                Pause(0.25)

                '20170403 LEE: Cannot control Adobe Acrobat from Visual Studio
                'Must use Word.export
                '         Dim strPrinter As String = GetPDFDriver()
                '         If Len(strPrinter) = 0 Then
                '             wd.ActiveDocument.ExportAsFixedFormat(OutputFileName:=strP, ExportFormat:=Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, _
                'OpenAfterExport:=False, OptimizeFor:=Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Range:= _
                'Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, Item:=Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, _
                'IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
                'Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
                'True, UseISO19005_1:=False)
                '         Else

                '             Try
                '                 Dim strPrinterOrig As String = wd.ActivePrinter.ToString

                '                 wd.ActivePrinter = strPrinter

                '                 wd.Application.PrintOut(FileName:=strP1, Range:=Microsoft.Office.Interop.Word.WdPrintOutRange.wdPrintAllDocument, Item:= _
                ' Microsoft.Office.Interop.Word.WdPrintOutItem.wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=Microsoft.Office.Interop.Word.WdPrintOutPages.wdPrintAllPages)

                '                 wd.ActivePrinter = strPrinterOrig
                '             Catch ex As Exception
                '                 var1 = ex.Message
                '                 var1 = var1
                '             End Try

                '         End If

                Try

                    wd.ActiveDocument.ExportAsFixedFormat(OutputFileName:=strP, ExportFormat:=Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, _
    OpenAfterExport:=False, OptimizeFor:=Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Range:= _
    Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, Item:=Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, _
    IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
    Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
    True, UseISO19005_1:=False)


                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try

                'wd.ActiveDocument.ExportAsFixedFormat(OutputFileName:=strP, ExportFormat:=Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF)

                strM = "The path for the created PDF file:" & ChrW(10) & ChrW(10) & strP & ChrW(10) & ChrW(10)
                strM1 = "PDF created..."
                CreatePDF = strP

                Pause(0.25)

                Try
                    System.Diagnostics.Process.Start(strP)
                Catch ex As Exception
                    strM = "There was problem opening this file as a PDF."
                    'strM = strM & ChrW(10) & ChrW(10) & "It may be that there is not a configured default PDF viewer on this workstation."
                    strM = strM & ChrW(10) & ChrW(10) & "It may be that this workstation does not have a configured default PDF viewer."
                    MsgBox(strM, MsgBoxStyle.Information, "Problem...")

                End Try
            End If

        Catch ex As Exception
            strM = "There was a problem creating PDF file:" & ChrW(10) & ChrW(10) & strP & ChrW(10) & ChrW(10)
            strM1 = strM & "PDF not created..." & ChrW(10) & ChrW(10)
            strM1 = strM1 & ex.Message

            MsgBox(strM1, MsgBoxStyle.Information, "PDF not created...")
            gDoPDF = False
        End Try

        var1 = 0 'debug

    End Function

    Function IsFile(ByVal strF As String) As Boolean

        IsFile = False

        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String



    End Function

    Function GetScNot(ByVal intDig As Short) As String

        GetScNot = "0.00000E0"

        GetScNot = "0.00000E0"
        'GetScNot = "0.00000E-0"

        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String

        If intDig = 1 Then
            GetScNot = "0E0"
            GoTo end1
        End If

        Try
            str1 = "0."
            For Count1 = 1 To intDig - 1
                str1 = str1 & "0"
            Next
            'str1 = str1 & "E-0"
            str1 = str1 & "E0"

            GetScNot = str1
        Catch ex As Exception

        End Try

end1:

    End Function

    Sub SetHighAndLowCriteria(ByVal varNom As Double, ByVal dblOffsetHigh As Double, ByVal dblOffsetLow As Double, ByRef hi As Double, ByRef lo As Double)
        'Given nominal Mass, and Percentage Offset value, this will fill in the values for hi and low
        'varNom = Nominal Mass,  Offset = (Percentage offset * 100)

        '20160820 LEE: Note that this function isn't needed anymore because the hi/lo concept has been replaced with function OutsideAccCrit
        'eventually we should remove the use of this function

        If gboolCritFullPrec Then
            hi = CDec(CDec(varNom) + (CDec(varNom) * CDec(dblOffsetHigh) / 100))
            lo = CDec(CDec(varNom) - (CDec(varNom) * CDec(dblOffsetLow) / 100))
        Else
            If boolLUseSigFigs Then
                hi = CDec(SigFigOrDec(CDec(varNom) + (CDec(varNom) * CDec(dblOffsetHigh) / 100), LSigFig, False))
                lo = CDec(SigFigOrDec(CDec(varNom) - (CDec(varNom) * CDec(dblOffsetLow) / 100), LSigFig, False))
            Else
                hi = CDec(RoundToDecimalRAFZ(CDec(varNom) + (CDec(varNom) * CDec(dblOffsetHigh) / 100), LSigFig))
                lo = CDec(RoundToDecimalRAFZ(CDec(varNom) - (CDec(varNom) * CDec(dblOffsetLow) / 100), LSigFig))
            End If
        End If

    End Sub

    Sub FixTableConcentrations(ByRef tbl As DataTable)

        Dim Count As Int32
        Dim intRoundToThisDecimal As Short
        intRoundToThisDecimal = 8
        For Count = 0 To tbl.Rows.Count - 1
            If (NZ(tbl.Rows.Item(Count).Item("Concentration"), False)) Then
                tbl.Rows.Item(Count).Item("Concentration") = RoundToDecimal(tbl.Rows.Item(Count).Item("Concentration"), 8)
            End If
        Next

    End Sub


    Function GetDECISIONREASONValue(boolFAS As Boolean, intAnalyteID As Int64, drAS As DataRow) As String

        'GetDECISIONREASONValue(boolExFromAS, vAnalyteID, drAS)
        'boolFAS:  true if from AssignedSamples

        GetDECISIONREASONValue = ""

        Dim strFAS As String
        Dim vAS1, vAS2, vAS3, vAS4

        If boolFAS Then
            GetDECISIONREASONValue = ""
        Else
            'need to get DECISIONREASON from tblBCQCConcs

            vAS1 = NZ(drAS.Item("RUNID"), 0)
            vAS2 = NZ(drAS.Item("RUNSAMPLESEQUENCENUMBER"), 0)

            strFAS = "ANALYTEID = " & intAnalyteID & " AND RUNID = " & vAS1 & " AND RUNSAMPLESEQUENCENUMBER = " & vAS2
            Dim rowsAS() As DataRow = tblBCQCConcs.Select(strFAS)
            If rowsAS.Length = 0 Then
                GetDECISIONREASONValue = ""
            Else
                vAS3 = rowsAS(0).Item("DECISIONREASON")
                GetDECISIONREASONValue = NZ(vAS3, "")
            End If
        End If

    End Function

    Function GetDECISIONREASON(var6 As Object, strO As String, idCT As Int64, boolParen As Boolean) As Object

        GetDECISIONREASON = strO

        Dim str2 As String

        If Len(NZ(var6, "")) = 0 Then
        Else
            str2 = Mid(var6, Len(var6), 1)
            If StrComp(str2, ".", CompareMethod.Text) = 0 Then
            Else
                var6 = var6 & "."
            End If
            var6 = var6 & " Value excluded from summary statistics."

            ''skip this and add to end of table instead
            If boolParen And boolQCREPORTACCVALUES = False Then
                Select Case idCT
                    Case 11, 32 'ANOVA table, AdHocStabilityComparison
                        If boolQCREPORTACCVALUES Then
                        Else
                            '20180419 LEE: Added 'statistical'
                            var6 = var6 & " The statistical results within parentheses were calculated including the outlier value."
                        End If
                    Case 17 'matrix effect
                        If BOOLDOINDREC Then
                            var6 = var6 & " The statistical results within parentheses were calculated including the outlier value."
                        End If
                End Select
            End If

            GetDECISIONREASON = var6
        End If

    End Function

    Function HasOutlier(ByVal strX As String, ByVal boolHasO As Boolean) As Boolean

        HasOutlier = False
        Dim str1 As String = "results within parentheses"

        If boolHasO Then
            HasOutlier = True
        Else
            If InStr(1, strX, str1, CompareMethod.Text) > 0 Then
                HasOutlier = True
            End If
        End If

    End Function

    Function GetLegendStringIncluded(ByVal hiPercent As Double, ByVal lowPercent As Double, ByVal boolUseGuwuCriteria As Int16) As String

        If gboolCritFullPrec Then
        Else
            hiPercent = RoundToDecimal(hiPercent, 0)
            lowPercent = RoundToDecimal(lowPercent, 0)
        End If

        If gidTR = 13 Or gidTR = 14 Or gidTR = 15 Then 'Recovery/MatrixFactor tables
            GetLegendStringIncluded = ""
        Else
            If gAllowGuWuAccCrit And LAllowGuWuAccCrit And boolUseGuwuCriteria = -1 Then
                If hiPercent = lowPercent Then
                    GetLegendStringIncluded = "Value outside of acceptance criteria (" & ChrW(177) & " " & hiPercent & " % theoretical)" ' but included in summary statistics."
                Else
                    GetLegendStringIncluded = "Value outside of acceptance criteria (+" & hiPercent & "/-" & lowPercent & " % theoretical)" ' but included in summary statistics."
                End If
            Else
                GetLegendStringIncluded = "Value outside of acceptance criteria (" & ChrW(177) & " " & hiPercent & "% theoretical)" ' but included in summary statistics."
            End If

            If boolSTATSMEAN Then
                GetLegendStringIncluded = GetLegendStringIncluded & " but included in summary statistics."
            End If
        End If


    End Function


    Function GetLegendStringExcluded(ByVal hiPercent As Double, ByVal lowPercent As Double, ByVal boolUseGuwuCriteria As Int16, varDR As Object, idCT As Int64, boolParen As Boolean, strDiff As String) As String

        Dim str1, str2 As String

        If gidTR = 13 Or gidTR = 14 Or gidTR = 15 Then 'Recovery/MatrixFactor tables
            str1 = "Excluded from summary statistics because the value is a statistical outlier according to the " & ReturnOutlierMethod() & "."
            GetLegendStringExcluded = str1
        Else
            str1 = GetLegendStringIncluded(hiPercent, lowPercent, boolUseGuwuCriteria)
            str2 = "and excluded from summary statistics because the value is a statistical outlier according to the " & ReturnOutlierMethod() & "."

            If boolParen And boolQCREPORTACCVALUES = False Then
                Select Case idCT
                    Case 11, 3, 32 'ANOVA table,CalibrTable, AdHocStabilityComparison
                        '20180419 LEE: Added 'statistical'
                        If Len(strDiff) > 0 Then
                            str2 = str2 & " The statistical results and " & strDiff & " values within parentheses were calculated including the outlier value."
                        Else
                            str2 = str2 & " The statistical results within parentheses were calculated including the outlier value."
                        End If

                    Case 17 'matrix effect
                        If BOOLDOINDREC Then
                            If Len(strDiff) > 0 Then
                                str2 = str2 & " The statistical results and " & strDiff & " values within parentheses were calculated including the outlier value."
                            Else
                                str2 = str2 & " The statistical results within parentheses were calculated including the outlier value."
                            End If
                        End If
                End Select
            Else
                If Len(strDiff) > 0 Then
                    str2 = str2 & " The " & strDiff & " values within parentheses were calculated including the outlier value."
                End If
            End If
            str1 = str1.Replace("but included in summary statistics.", str2)
            '20151002 Larry: Added by Larry
            GetLegendStringExcluded = str1
        End If

        If Len(NZ(varDR, "")) = 0 Then
        Else
            GetLegendStringExcluded = GetDECISIONREASON(varDR, str1, idCT, boolParen)
        End If

    End Function

    Function GetLegendStringExcludedRegression(ByVal hiPercent As Double, ByVal lowPercent As Double, ByVal boolUseGuwuCriteria As Int16) As String

        Dim str1, str2 As String
        str1 = GetLegendStringIncluded(hiPercent, lowPercent, boolUseGuwuCriteria)
        str2 = "and excluded from regression and summary statistics." & ReturnOutlierMethod() & "."
        str1 = str1.Replace("but included in summary statistics.", str2)
        str1 = "Not Reported: " & str1
        '20151002 Larry: Added by Larry
        GetLegendStringExcludedRegression = str1

    End Function

    Function SetLegendArray(ByRef arrLegend As Object, ByRef intLegLineNum As Int16, ByVal strLegendLine As String, ByRef strLetterReference As String, boolSuper As Boolean) As Int16

        'Returns number of items added to the legend (1 or 0), updates intLegLineNum and strLetterReference if the line is already in the legend.

        '20180201 LEE:
        '1= Actual string to search in table
        '2= Not used in SplitTable
        '3= True/False to superscript table
        '4= True: Do not look for item in table, but add buffer row to row count.  False: Look for item in table; if found, add buffer row to row count


        Dim intItemsAdded As Int16 = 0
        Dim boolPro As Boolean
        Dim Count1 As Int16
        Dim str2 As String

        If intLegLineNum = 1 Then 'First Legend Entry
            arrLegend(1, intLegLineNum) = strLetterReference
            arrLegend(2, intLegLineNum) = strLegendLine
            arrLegend(3, intLegLineNum) = boolSuper 'True
            intItemsAdded = 1
        Else
            boolPro = True
            For Count1 = 1 To intLegLineNum - 1 'Check previous entries for a duplicate
                str2 = arrLegend(2, Count1)
                If StrComp(strLegendLine, str2, CompareMethod.Text) = 0 Then 'Same as previous legend line; abort.
                    intLegLineNum = intLegLineNum - 1
                    strLetterReference = arrLegend(1, Count1)
                    boolPro = False
                    Exit For

                End If
            Next
            If boolPro Then
                arrLegend(1, intLegLineNum) = strLetterReference
                arrLegend(2, intLegLineNum) = strLegendLine
                arrLegend(3, intLegLineNum) = boolSuper ' True
                intItemsAdded = intItemsAdded + 1
            End If
        End If
        SetLegendArray = intItemsAdded
    End Function

    Function CalcREPercent(ByVal numMean As Decimal, ByVal varNom As Decimal, ByVal intQCDec As Short) As Decimal
        'numMean = average for the data
        'varNom = nominal concentration
        'intQCDec = number of decimal places to round to

        '20180430 LEE:
        'For Endogenous Cmpds, NomConc = 0
        If varNom <= 0 Then
            CalcREPercent = 0
        Else
            If CDec(varNom) = 0 Then
                CalcREPercent = 0
            Else
                CalcREPercent = RoundToDecimal(RoundToDecimalRAFZ(((numMean / CDec(varNom)) - 1) * 100, intQCDec + 4), intQCDec)
            End If
        End If


    End Function

    Function CalcCVPercent(ByVal numSD As Decimal, ByVal numMean As Decimal, ByVal intQCDec As Short) As Decimal
        'numSD = standard deviation
        'numMean = average for the data
        'intQCDec = number of decimal places to round to
        If numMean = 0 Then
            CalcCVPercent = 0
        Else
            CalcCVPercent = RoundToDecimal(RoundToDecimalRAFZ((numSD / numMean * 100), intQCDec + 4), intQCDec)
        End If

    End Function


    Function ConvertVECandNM(ByVal str As String) As String
        If (StrComp(str, "VEC") = 0) Then
            ConvertVECandNM = AQL()
        ElseIf (StrComp(str, "NM") = 0) Then
            ConvertVECandNM = BQL()
        Else
            ConvertVECandNM = str
        End If
    End Function

    Function EnforceMinimumTableVerticalSpacing(ByRef wd As Microsoft.Office.Interop.Word.Application, intMinSpacing As Short) As Boolean

        Dim spaceAfter As Short
        EnforceMinimumTableVerticalSpacing = False
        If ((wd.Selection.ParagraphFormat.SpaceAfter + wd.Selection.ParagraphFormat.SpaceBefore) < intMinSpacing) Then
            'Increase the space after to ensure minimal spacing
            spaceAfter = intMinSpacing - wd.Selection.ParagraphFormat.SpaceBefore
            If (spaceAfter >= 0) Then  'Being slightly paranoid here (would only be true if SpaceAfter was negative)
                wd.Selection.ParagraphFormat.SpaceAfter = spaceAfter
                EnforceMinimumTableVerticalSpacing = True
            End If
        End If

    End Function

    Function atLeastWordVersion10(ByRef wd As Microsoft.Office.Interop.Word.Application)
        If CInt(wd.Version) >= 14 Then
            atLeastWordVersion10 = True
        Else
            atLeastWordVersion10 = False
        End If
    End Function


    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer

        'http://stackoverflow.com/questions/5193893/count-specific-character-occurrences-in-string

        Return Len(value) - Len(Replace(value, ch, ""))

    End Function

    Function getRunAnalyteLLOQ(ByVal strRunID As String, ByVal analyteID As Int64, intGroup As Int16) As Single

        getRunAnalyteLLOQ = 0

        'Gets LLOQ for Run & AnalyteID
        'NDL Being super-cautious here - otherwise we could do it in 3 lines
        '20160310 LEE: If sample value comes from Mean, then runID is string
        'changed runID parameter to string

        '20180629 LEE:
        'Note that intGroup can't be used here
        'This function is being called from Sample Concetrations table
        'This table may have multiple intGroups because all C_1, C_2 are reported in this one table

        Dim runID As Short
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim intN As Short
        Dim str1 As String

        intN = CountCharacter(strRunID, ",") 'will return 0 if no commas

        If intN = 0 Then
            runID = CInt(strRunID)
            getRunAnalyteLLOQ = getRunAnalyteLimit("LLOQ", runID, analyteID, intGroup)
            If (IsNothing(getRunAnalyteLLOQ)) Then
                MsgBox("getRunAnalyteLLOQ: Could not find LLOQ in tblCalStdGroupAssayIDsAcc.")
            End If
        Else
            'get runids from strrunid
            'this code will get lloq for 1st runid involved in Mean
            int1 = 1
            int2 = 1
            For Count1 = 1 To intN
                int2 = InStr(int1, strRunID, ",", CompareMethod.Text)
                If Count1 = 1 Then
                    str1 = Mid(strRunID, int1, int2 - int1)
                    If IsNumeric(str1) Then
                        runID = CInt(str1)
                        getRunAnalyteLLOQ = getRunAnalyteLimit("LLOQ", runID, analyteID, intGroup)
                        If (IsNothing(getRunAnalyteLLOQ)) Then
                            MsgBox("getRunAnalyteLLOQ: Could not find LLOQ in tblCalStdGroupAssayIDsAcc.")
                        End If
                        Exit For
                    End If
                End If
            Next
        End If


    End Function

    Function getRunAnalyteULOQ(ByVal strRunID As String, ByVal analyteID As Int64, intGroup As Int16) As Single

        getRunAnalyteULOQ = 0

        'NDL Gets ULOQ for Run & AnalyteID
        'Being super-cautious here - otherwise we could do it in 3 lines

        '20160310 LEE: If sample value comes from Mean, then runID is string
        'changed runID parameter to string

        '20180629 LEE:
        'Note that intGroup can't be used here
        'This function is being called from Sample Concetrations table
        'This table may have multiple intGroups because all C_1, C_2 are reported in this one table


        Dim runID As Short
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim intN As Short
        Dim str1 As String

        intN = CountCharacter(strRunID, ",") 'will return 0 if no commas

        If intN = 0 Then
            runID = CInt(strRunID)
            getRunAnalyteULOQ = getRunAnalyteLimit("ULOQ", runID, analyteID, intGroup)
            If (IsNothing(getRunAnalyteULOQ)) Then
                MsgBox("getRunAnalyteULOQ: Could not find LLOQ in tblCalStdGroupAssayIDsAcc.")
            End If
        Else
            'get runids from strrunid
            'this code will get lloq for 1st runid involved in Mean

            int1 = 1
            int2 = 1
            For Count1 = 1 To intN
                int2 = InStr(int1, strRunID, ",", CompareMethod.Text)
                If Count1 = 1 Then
                    str1 = Mid(strRunID, int1, int2 - int1)
                    If IsNumeric(str1) Then
                        runID = CInt(str1)
                        getRunAnalyteULOQ = getRunAnalyteLimit("ULOQ", runID, analyteID, intGroup)
                        If (IsNothing(getRunAnalyteULOQ)) Then
                            MsgBox("getRunAnalyteULOQ: Could not find LLOQ in tblCalStdGroupAssayIDsAcc.")
                        End If
                        Exit For
                    End If
                End If
            Next
        End If

    End Function

    Function getRunAnalyteLimit(ByVal strLimit As String, ByVal runID As Integer, ByVal analyteID As Int64, intGroup As Int16) As String

        'Gets "ULOQ" or "LLOQ" for Run & AnalyteID
        Dim dr() As DataRow
        Dim strF As String
        'Initialize
        If (StrComp(strLimit, "LLOQ") = 0) Then
            getRunAnalyteLimit = "10000000"  'Set LLOQ to high number so it will trigger if something goes wrong
        ElseIf (StrComp(strLimit, "ULOQ") = 0) Then
            getRunAnalyteLimit = "-1"  'Set ULOQ to low number so it will trigger if something goes wrong
        Else
            MsgBox("getRunAnalyteLimit: Wrong Input")
        End If

        '20180629 LEE:
        'Note that intGroup can't be used here
        'This function is being called from Sample Concetrations table
        'This table may have multiple intGroups because all C_1, C_2 are reported in this one table


        'strF = "RUNID = '" & runID & "' AND ANALYTEID = '" & analyteID & "' AND INTGROUP = " & intGroup
        strF = "RUNID = '" & runID & "' AND ANALYTEID = '" & analyteID & "'"
        dr = tblCalStdGroupAssayIDsAcc.Select(strF)
        If (dr.Length < 1) Then
            Console.Write("In getRunAnalyteLimit(" & strLimit & "): No Analyte Groups found for runID " & runID & "and AnalyteID " & analyteID & ".")
        ElseIf (dr.Length > 1) Then
            Console.Write("In getRunAnalyte(" & strLimit & "): Multiple Analyte Groups found for runID " & runID & "and AnalyteID " & analyteID & ".")
        Else
            If (StrComp(strLimit, "LLOQ") = 0) Then
                getRunAnalyteLimit = dr(0).Item("LLOQ")
            ElseIf (StrComp(strLimit, "ULOQ") = 0) Then
                getRunAnalyteLimit = dr(0).Item("ULOQ")
            End If
        End If

    End Function

    Function numTablesToBeGenerated(ByVal intTableNumber As Short, ByVal boolTableForEachMatrix As Boolean, ByVal boolTableForEachRange As Boolean)

        Dim strMatrix, strAnalyteID, strDo As String
        Dim CountAnalyte, CountMatrix, CountSubAnalytes As Short
        Dim dv As New DataView(tblAnalyteGroups)
        Dim dvDo As System.Data.DataView

        'This function simply counts the number of tables which are to be generated.
        'The number is used in the Status updates.

        numTablesToBeGenerated = 0
        dvDo = frmH.dgvReportTableConfiguration.DataSource

        For CountAnalyte = 0 To tblAnalyteIDs.Rows.Count - 1 'For each AnalyteID
            strAnalyteID = tblAnalyteIDs.Rows(CountAnalyte).Item("AnalyteID")
            dv.RowFilter = "ANALYTEID = " & strAnalyteID
            If (Not boolTableForEachMatrix) And (Not boolTableForEachRange) Then
                'Just count the analyteID's
                For CountSubAnalytes = 0 To dv.Count
                    strDo = dv(CountSubAnalytes).Item("ANALYTEDESCRIPTION_C") 'The Sub-Analyte name (eg. CmpdXYZ_C1)
                    If dvDo.Item(intTableNumber).Item(strDo) Then 'If analyte is checked *for this table*
                        numTablesToBeGenerated = numTablesToBeGenerated + 1
                        Exit For 'One is enough for each AnalyteID
                    End If
                Next

            Else 'Count AnalyteID's and Matrices
                For CountMatrix = 0 To tblMatrices.Rows.Count - 1
                    strMatrix = tblMatrices.Rows(CountMatrix).Item("Matrix")
                    dv.RowFilter = "ANALYTEID = " & strAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                    For CountSubAnalytes = 0 To dv.Count - 1
                        strDo = dv(CountSubAnalytes).Item("ANALYTEDESCRIPTION_C") 'The Sub-Analyte name (eg. CmpdXYZ_C1)
                        If dvDo.Item(intTableNumber).Item(strDo) Then 'If analyte is checked *for this table*
                            numTablesToBeGenerated = numTablesToBeGenerated + 1
                            If (Not boolTableForEachRange) Then 'One is enough for each Matrix
                                Exit For 'Go to next Matrix
                            Else 'Count every subAnalyte
                            End If
                        End If
                    Next 'SubAnalyte
                Next 'Matrix
            End If
        Next 'Analyte
    End Function
End Module
