Option Compare Text

Module modMethValFill

    Public tblInterQCSum As New System.Data.DataTable

    Sub ValSummaryTable(ByVal wd As Microsoft.Office.Interop.Word.Application)



    End Sub

    Function GetRecovery(ByVal dtbl As System.Data.DataTable, ByVal idS As Int64, ByVal idT As Int64, ByVal strAnal As String, ByVal boolIS As Boolean, ByVal strIS As String) As String

        GetRecovery = "NA"

        Exit Function

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
        Dim tbl2a As New System.Data.DataTable
        Dim tbl2b As New System.Data.DataTable
        Dim dv2 As System.Data.DataView
        Dim rows2() As DataRow
        Dim rows2a() As DataRow
        Dim rows2b() As DataRow
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
        Dim arrFP(20) 'FlagPercent array
        Dim strFP As String
        Dim numMean As Decimal
        Dim numBias As Decimal
        Dim numSD As Decimal
        Dim tblZ As New System.Data.DataTable
        Dim dvAn As System.Data.DataView
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
        Dim strX As String
        Dim rowsX() As DataRow
        Dim boolX As Boolean
        Dim intCols As Short
        Dim col1, col2, col3, col4, col5, col6 As Short
        Dim intLegStart As Short
        Dim intRowsData As Short
        Dim boolHasNA As Boolean
        Dim boolPro As Boolean
        Dim boolJustTable As Boolean
        Dim numNomConc As Decimal

        Dim intExp As Short
        Dim ctExp As Short
        Dim int8 As Short

        Dim rows1E() As DataRow
        Dim rows3E() As DataRow
        Dim nE As Short
        Dim nI As Short
        Dim boolOutHeadE As Boolean = False
        Dim boolOutHeadI As Boolean = False
        Dim boolDeleteRows As Boolean = False
        Dim boolE() As Boolean
        Dim boolETot As Boolean
        Dim intStart As Short
        Dim numA, numB, numAI, numBI
        Dim intRowCol6 As Short
        Dim intRowCol6I As Short
        Dim arr1(2, 1)
        Dim arr2(1)

        Try

            boolSTATSBIAS = False 'bias not possible in this table
            boolTHEORETICAL = False 'theoretical not possible in this table

            boolJustTable = False

            boolHasNA = False

            strDo = strAnal

            ''''wdd.visible = True

            intTableID = idT

            dvDo = frmH.dgvReportTableConfiguration.DataSource
            strF = "id_tblconfigreporttables = " & intTableID
            intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")
            intLeg = 0
            intLegStart = 96
            boolPro = False

            ''Get table name
            'var1 = dvDo(intDo).Item("Table")
            'strTName = NZ(var1, "[NONE]")

            ''get Temperature info
            'var1 = dvDo(intDo).Item("PERIODTEMP")
            'strTempInfo = NZ(var1, "[NONE]")

            '***
            'intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")
            'intLeg = 0
            'intLegStart = 96
            'boolPro = False

            ''Get table name
            ''var1 = dvDo(intDo).Item("Table")
            'var1 = dvDo(intDo).Item("CHARHEADINGTEXT")
            'strTName = NZ(var1, "[NONE]")

            ''get Temperature info
            'var1 = dvDo(intDo).Item("CHARSTABILITYPERIOD")
            'strTempInfo = NZ(var1, "[NONE]")

            '***
            tbl1 = tblAnalysisResultsHome
            tbl2 = tblAssignedSamples
            tbl3 = tblAssignedSamplesHelper
            tbl4 = tblAnalytesHome

            'ensure data has been entered
            strF = "id_tblconfigreporttables = " & intTableID & " AND id_tblStudies = " & idS ' & " AND ID_TBLREPORTTABLE = " & idTR
            rowsX = tbl2.Select(strF)
            If rowsX.Length = 0 Then
                Exit Function
            End If

            strF = "IsIntStd = 'No' OR IsIntStd = 'Yes'"
            strS = "IsIntStd ASC, AnalyteDescription ASC"
            rows11 = tblAnalytesHome.Select(strF, strS)
            intRowsAnal = rows11.Length

            intStart = 1
            For Count1 = 0 To intRowsAnal - 1
                str1 = rows11(Count1).Item("ANALYTEDESCRIPTION")
                If StrComp(str1, strAnal, CompareMethod.Text) = 0 Then
                    intStart = Count1 + 1
                    Exit For
                End If

            Next

            'For Count1 = 1 To intRowsAnal

            For Count1 = intStart To intStart

                Dim arrLegend(4, 20)

                'check if table is to be generated
                'strDo = arrAnalytes(1, Count1) 'record column name
                strDo = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")
                strX = rows11(Count1 - 1).Item("IsIntStd")
                boolX = False
                boolJustTable = False
                If StrComp(strX, "Yes", CompareMethod.Text) = 0 Then
                    'check for boolIntStd in tbl2
                    strF = "IsIntStd = 'Yes'"
                    var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                    var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                    var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                    strF2 = "ID_TBLSTUDIES = " & idS & " AND "
                    strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                    'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                    strF2 = strF2 & "BOOLINTSTD = -1"
                    Erase rowsX
                    rowsX = tbl2.Select(strF2)
                    int1 = rowsX.Length
                    If int1 > 0 Then
                        bool = True
                        boolX = True
                    Else
                        bool = True 'False
                    End If
                Else
                    'bool = dvDo.Item(intDo).Item(strDo) 'find boolean value of dvDo column
                End If
                bool = True


                If bool Then 'continue
                    'ensure data has been entered
                    If StrComp(strX, "Yes", CompareMethod.Text) = 0 Then
                        strF = strF2
                    Else
                        strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & idS & " AND CHARANALYTE = '" & CleanText(strDo) & "'" ' AND ID_TBLREPORTTABLE = " & idTR

                    End If
                    rowsX = tbl2.Select(strF)
                    If rowsX.Length = 0 Then
                        boolJustTable = True
                        GoTo end1
                    Else
                        boolJustTable = False
                    End If


                    ''get strConcUnits
                    'int1 = FindRowDV("ULOQ Units", frmH.dgvWatsonAnalRef.DataSource)
                    'strConcUnits = NZ(frmH.dgvWatsonAnalRef(Count1, int1).Value, "ng/mL")

                    'setup tables
                    If boolX Then
                        var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        strF2 = "BOOLINTSTD = -1 AND "
                        strF2 = strF2 & "ID_TBLSTUDIES = " & idS & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
                    Else
                        strF2 = "BOOLINTSTD = 0 AND "
                        var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                        var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                        var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                        strF2 = strF2 & "ID_TBLSTUDIES = " & idS & " AND "
                        strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                        'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                        strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                        strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "
                    End If
                    strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
                    rows2 = tbl2.Select(strF2, strS)
                    int1 = rows2.Length 'debug
                    dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                    int1 = dv2.Count 'debug

                    'find number of runs used
                    tblNumRuns = dv2.ToTable("a", True, "RUNID")
                    intNumRuns = tblNumRuns.Rows.Count

                    'establish number of QCs evaluated
                    'this will actually give number of rows
                    Dim dvSC As System.Data.DataView = New DataView(tbl2, strF2, "NOMCONC ASC", DataViewRowState.CurrentRows)
                    tblLevels = dvSC.ToTable("SC", True, "NOMCONC")
                    intNumLevels = tblLevels.Rows.Count
                    For Count2 = 0 To intNumLevels - 1 'check for any null values
                        var3 = tblLevels.Rows.Item(Count2).Item("NOMCONC")
                        If IsDBNull(var3) Then
                            'str1 = "Dude, the Nominal Concentration for some assigned samples for " & rows11(Count1 - 1).Item("ANALYTEDESCRIPTION") & " have not been configured."
                            'str1 = str1 & ChrW(10) & "When this action is finished, please navigate to the Assigned Samples window and correct this problem."
                            'MsgBox(str1, MsgBoxStyle.Information, "Nom Conc problem...")
                            GoTo end1
                        End If
                    Next

                    ReDim arr1(2, intNumLevels)
                    ReDim boolE(intNumLevels)

                    'find number of table rows to generate
                    intRowsX = intNumLevels

                    'find introwsQC and dataviews CHARHELPER1
                    If boolX Then
                        strF = " AND CHARHELPER1 = 'QC' AND BOOLINTSTD = -1"
                    Else
                        strF = " AND CHARHELPER1 = 'QC' AND BOOLINTSTD = 0"
                    End If
                    strF = strF2 & strF
                    strS = "NOMCONC ASC, RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
                    dv1 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tbl2a = dv1.ToTable
                    introwsQC = tbl2a.Rows.Count

                    'find introwsRS and dataviews CHARHELPER1
                    If boolX Then
                        strF = " AND CHARHELPER1 = 'PES - Post Extraction Spike' AND BOOLINTSTD = -1"
                    Else
                        strF = " AND CHARHELPER1 = 'PES - Post Extraction Spike' AND BOOLINTSTD = 0"
                    End If

                    'strF = " AND CHARHELPER1 = 'PES - Post Extraction Spike'"
                    strF = strF2 & strF
                    strS = "NOMCONC ASC, RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
                    dv3 = New DataView(tbl2, strF, strS, DataViewRowState.CurrentRows)
                    tbl2b = dv3.ToTable
                    introwsRS = tbl2b.Rows.Count

                    If introwsRS > introwsQC Then
                        intRowsData = introwsRS
                    Else
                        intRowsData = introwsQC
                    End If

                    'first determine if there are any outliers
                    Dim intInc As Short
                    Dim intRowsT As Short
                    Dim RowsT() As DataRow
                    boolETot = False
                    intExp = 0
                    If boolQCREPORTACCVALUES Then
                    Else
                        For Count2 = 0 To intNumRuns - 1
                            var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")
                            intInc = 0
                            For Count3 = 0 To intNumLevels - 1
                                varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                                boolE(Count3) = False

                                'column 2
                                Select Case Count3
                                    Case 0
                                        str1 = "QC Low"
                                    Case 1
                                        str1 = "QC Mid"
                                    Case 2
                                        str1 = "QC High"
                                End Select

                                'column 4
                                'get average of peak areas for QC
                                'and record column 4 peak areas for QC
                                strF = "NOMCONC = " & varNom
                                Erase rows2a
                                rows2a = tbl2a.Select(strF)
                                int2 = rows2a.Length
                                var1 = rows2a(0).Item("BOOLINTSTD")
                                If var1 = -1 Then
                                    boolIS = True
                                Else
                                    boolIS = False
                                End If
                                strF = ""
                                For Count4 = 0 To int2 - 1
                                    var2 = rows2a(Count4).Item("ANALYTEINDEX")
                                    var3 = rows2a(Count4).Item("MASTERASSAYID")
                                    var4 = rows2a(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                                    If Count4 <> int2 - 1 Then
                                        strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ") OR "
                                    Else
                                        strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ")"
                                    End If
                                Next
                                Erase rows1
                                rows1 = tbl1.Select(strF)
                                int3 = rows1.Length
                                Erase RowsT
                                If introwsRS > introwsQC Then
                                    strF = "NOMCONC = " & varNom
                                    Erase rows2b
                                    RowsT = tbl2b.Select(strF)
                                    intRowsT = RowsT.Length
                                Else
                                    intRowsT = int3
                                End If

                                For Count4 = 0 To intRowsT - 1
                                    If Count4 > int3 - 1 Then
                                        var1 = "NA"
                                        boolHasNA = True 'for legend purposes
                                    Else
                                        If boolIS Then
                                            var1 = rows1(Count4).Item("INTERNALSTANDARDAREA")
                                        Else
                                            var1 = rows1(Count4).Item("ANALYTEAREA")
                                        End If
                                    End If

                                    If Count4 > int3 - 1 Then
                                    Else
                                        var2 = NZ(rows1(Count4).Item("ELIMINATEDFLAG"), "N")
                                        'check for outlier
                                        If StrComp(var2, "Y", vbTextCompare) = 0 Then
                                            boolE(Count3) = True
                                            boolETot = True
                                            Exit For
                                        End If
                                    End If
                                Next Count4

                                If boolE(Count3) Then
                                Else
                                    'get column 5
                                    'get average of peak areas for RS
                                    strF = "NOMCONC = " & varNom
                                    Erase rows2b
                                    rows2b = tbl2b.Select(strF)
                                    int2 = rows2b.Length

                                    var1 = rows2b(0).Item("BOOLINTSTD")
                                    If var1 = -1 Then
                                        boolIS = True
                                    Else
                                        boolIS = False
                                    End If
                                    strF = ""
                                    For Count4 = 0 To int2 - 1
                                        var2 = rows2b(Count4).Item("ANALYTEINDEX")
                                        var3 = rows2b(Count4).Item("MASTERASSAYID")
                                        var4 = rows2b(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                                        If Count4 <> int2 - 1 Then
                                            strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ") OR "
                                        Else
                                            strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ")"
                                        End If
                                    Next
                                    Erase rows3
                                    rows3 = tbl1.Select(strF)
                                    int3 = rows3.Length 'debug

                                    For Count4 = 0 To intRowsT - 1
                                        If Count4 > int3 - 1 Then
                                            var1 = "NA"
                                            boolHasNA = True 'for legend purposes
                                        Else
                                            If boolIS Then
                                                var1 = rows3(Count4).Item("INTERNALSTANDARDAREA")
                                            Else
                                                var1 = rows3(Count4).Item("ANALYTEAREA")
                                            End If
                                        End If

                                        If Count4 > int3 - 1 Then
                                        Else
                                            var2 = NZ(rows3(Count4).Item("ELIMINATEDFLAG"), "N")
                                            'check for outlier
                                            If StrComp(var2, "Y", vbTextCompare) = 0 Then
                                                intExp = intExp + 1
                                                boolE(Count3) = True
                                                Exit For
                                            End If
                                        End If
                                    Next Count4
                                End If
                            Next Count3
                        Next Count2
                    End If


                    'begin entering data'
                    int1 = 3 'row position counter
                    For Count2 = 0 To intNumRuns - 1

                        'enter runid
                        var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")

                        'start filling in data by rows
                        'intRowsX = 0

                        intInc = 0
                        For Count3 = 0 To intNumLevels - 1

                            varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")

                            'column 2
                            Select Case Count3
                                Case 0
                                    str1 = "QC Low"
                                Case 1
                                    str1 = "QC Mid"
                                Case 2
                                    str1 = "QC High"
                            End Select
                            '.Selection.Tables.Item(1).Cell(int1 + intInc, 2).Select()
                            '.Selection.TypeText(str1)
                            If boolX Then
                            Else
                                'column 3
                                '.Selection.Tables.Item(1).Cell(int1 + intInc, col3).Select()
                                '.Selection.TypeText(CStr(SigFigOrDecString(varNom, LSigFig, False)))
                            End If

                            'column 4
                            'get average of peak areas for QC
                            'and record column 4 peak areas for QC
                            strF = "NOMCONC = " & varNom
                            Erase rows2a
                            rows2a = tbl2a.Select(strF)
                            int2 = rows2a.Length
                            var1 = rows2a(0).Item("BOOLINTSTD")
                            If var1 = -1 Then
                                boolIS = True
                            Else
                                boolIS = False
                            End If
                            strF = ""
                            For Count4 = 0 To int2 - 1
                                var2 = rows2a(Count4).Item("ANALYTEINDEX")
                                var3 = rows2a(Count4).Item("MASTERASSAYID")
                                var4 = rows2a(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                                If Count4 <> int2 - 1 Then
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ") OR "
                                Else
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ")"
                                End If
                            Next
                            Erase rows1
                            rows1 = tbl1.Select(strF)
                            int3 = rows1.Length
                            nI = int3

                            'now do rowsExcluded
                            strF = ""
                            For Count4 = 0 To int2 - 1
                                var2 = rows2a(Count4).Item("ANALYTEINDEX")
                                var3 = rows2a(Count4).Item("MASTERASSAYID")
                                var4 = rows2a(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                                If Count4 <> int2 - 1 Then
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND (ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL)) OR "
                                Else
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND (ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL))"
                                End If
                            Next

                            Erase rows1E
                            rows1E = tbl1.Select(strF)
                            nE = rows1E.Length

                            Erase RowsT
                            If introwsRS > introwsQC Then
                                strF = "NOMCONC = " & varNom
                                Erase rows2b
                                RowsT = tbl2b.Select(strF)
                                intRowsT = RowsT.Length
                            Else
                                intRowsT = int3
                            End If

                            'For Count4 = 0 To intRowsT - 1
                            '    If Count4 > int3 - 1 Then
                            '        var1 = "NA"
                            '        boolHasNA = True 'for legend purposes
                            '    Else
                            '        If boolIS Then
                            '            var1 = rows1(Count4).Item("INTERNALSTANDARDAREA")
                            '        Else
                            '            var1 = rows1(Count4).Item("ANALYTEAREA")
                            '        End If
                            '    End If

                            '    '.Selection.Tables.Item(1).Cell(int1 + intInc + Count4, col4).Select()
                            '    '.Selection.TypeText(CStr(var1))
                            '    ''''''''''''''console.writeline(CStr(rows1(Count4).Item("RUNSAMPLESEQUENCENUMBER")) & ", " & var1)

                            '    If Count4 > int3 - 1 Then
                            '    Else
                            '        'var2 = rows1(Count4).Item("ELIMINATEDFLAG")
                            '        ''check for outlier
                            '        'If StrComp(var2, "Y", vbTextCompare) = 0 Then
                            '        '    intExp = intExp + 1
                            '        '    intLeg = intLeg + 1
                            '        '    strA = ChrW(intLeg + intLegStart)
                            '        '    str1 = "Value excluded from summary statistics because it is a statistical outlier according to the [OUTLIERMETHOD]."
                            '        '    'search for str1 in arrLegend
                            '        '    If intLeg = 1 Then
                            '        '        arrLegend(1, intLeg) = strA
                            '        '        arrLegend(2, intLeg) = str1
                            '        '        arrLegend(3, intLeg) = True
                            '        '        ctLegend = ctLegend + 1
                            '        '    Else
                            '        '        boolPro = True
                            '        '        For Count5 = 1 To intLeg - 1
                            '        '            str2 = arrLegend(2, Count5)
                            '        '            If StrComp(str1, str2, CompareMethod.Text) = 0 Then 'abort
                            '        '                intLeg = intLeg - 1
                            '        '                strA = arrLegend(1, Count5)
                            '        '                boolPro = False
                            '        '                Exit For
                            '        '            End If
                            '        '        Next
                            '        '        If boolPro Then
                            '        '            arrLegend(1, intLeg) = strA
                            '        '            arrLegend(2, intLeg) = str1
                            '        '            arrLegend(3, intLeg) = True
                            '        '            ctLegend = ctLegend + 1
                            '        '        End If
                            '        '    End If
                            '        '    .Selection.Font.Bold = True
                            '        '    .Selection.Font.Color = Word.WdColor.wdColorRed
                            '        '    .Selection.Font.Superscript = True
                            '        '    .Selection.Font.Size = 12
                            '        '    .Selection.TypeText(" " & strA)
                            '        '    .Selection.Font.Superscript = False
                            '        '    .Selection.Font.Size = fontsize


                            '        'End If

                            '    End If
                            'Next

                            'now dow Means, SD, etc
                            'enter titles
                            int2 = -1


                            intStart = int2
                            intRowCol6 = int1 + intInc + intRowsT + 1 + int2 + 1

                            boolSTATSMEAN = True
                            boolSTATSSD = True
                            boolSTATSCV = True
                            boolSTATSBIAS = True
                            boolSTATSN = True

                            boolQCREPORTACCVALUES = True

                            int2 = intStart
                            If boolSTATSMEAN Then
                                Try
                                    'do mean
                                    If boolIS Then
                                        var1 = MeanDR(rows1E, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                                    Else
                                        var1 = MeanDR(rows1E, "ANALYTEAREA", False, "gaga", False, False)
                                    End If
                                    int2 = int2 + 1
                                    var2 = RoundToDecimal(var1, 0)
                                    '.Selection.TypeText(CStr(var2))
                                    numA = var2
                                    numMean = numA
                                Catch ex As Exception

                                End Try

                            End If
                            'If boolSTATSSD Then
                            '    Try
                            '        'do SD
                            '        int2 = int2 + 1
                            '        If int3 < 3 Then
                            '            var2 = "NA"
                            '            boolHasNA = True 'for legend purposes
                            '        Else
                            '            If boolIS Then
                            '                var1 = StdDevDR(rows1E, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                            '            Else
                            '                var1 = StdDevDR(rows1E, "ANALYTEAREA", False, "gaga", False, False)
                            '            End If
                            '            var2 = RoundToDecimal(var1, 0)
                            '            numSD = var2
                            '        End If
                            '        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '        '.Selection.TypeText(CStr(var2))
                            '    Catch ex As Exception

                            '    End Try

                            'End If
                            'If boolSTATSCV Then
                            '    Try
                            '        'do %CV: numSD / numMean
                            '        int2 = int2 + 1
                            '        If int3 < 3 Then
                            '            var2 = "NA"
                            '            boolHasNA = True 'for legend purposes
                            '        Else
                            '            var1 = (numSD / numMean) * 100
                            '            var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '        End If
                            '        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '        '.Selection.TypeText(var2)
                            '    Catch ex As Exception

                            '    End Try

                            'End If
                            'If boolSTATSBIAS and boolstatsmean  Then
                            '    'do %Bias: Mean/Nominal
                            '    Try
                            '        int2 = int2 + 1
                            '        If int3 < 3 Then
                            '            var2 = "NA"
                            '            boolHasNA = True 'for legend purposes
                            '        Else

                            '            Try
                            '                var1 = ((numMean / numNomConc) - 1) * 100
                            '                var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '            Catch ex As Exception
                            '                var2 = "NA"
                            '                boolHasNA = True 'for legend purposes
                            '            End Try

                            '        End If
                            '        '.Selection.TypeText(var2)
                            '    Catch ex As Exception

                            '    End Try

                            'End If
                            'If boolSTATSN Then
                            '    Try
                            '        'n
                            '        int2 = int2 + 1
                            '        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '        'var2 = RoundToDecimal(var1, 0)
                            '        '.Selection.TypeText(CStr(int3))
                            '        '.Selection.TypeText(CStr(nE))

                            '    Catch ex As Exception

                            '    End Try
                            'End If

                            'If boolQCREPORTACCVALUES Then
                            'Else
                            '    If boolE(Count3) Then
                            '        int2 = int2 + 2

                            '        If boolSTATSMEAN Then
                            '            Try
                            '                'do mean
                            '                If boolIS Then
                            '                    var1 = MeanDR(rows1, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                            '                Else
                            '                    var1 = MeanDR(rows1, "ANALYTEAREA", False, "gaga", False, False)
                            '                End If
                            '                int2 = int2 + 1
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '                var2 = RoundToDecimal(var1, 0)
                            '                '.Selection.TypeText(CStr(var2))
                            '                numAI = var2
                            '                numMean = numAI
                            '            Catch ex As Exception

                            '            End Try

                            '        End If
                            '        If boolSTATSSD Then
                            '            Try
                            '                'do SD
                            '                int2 = int2 + 1
                            '                If int3 < 3 Then
                            '                    var2 = "NA"
                            '                    boolHasNA = True 'for legend purposes
                            '                Else
                            '                    If boolIS Then
                            '                        var1 = StdDevDR(rows1, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                            '                    Else
                            '                        var1 = StdDevDR(rows1, "ANALYTEAREA", False, "gaga", False, False)
                            '                    End If
                            '                    var2 = RoundToDecimal(var1, 0)
                            '                    numSD = var2
                            '                End If
                            '                ' .Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '                '.Selection.TypeText(CStr(var2))
                            '            Catch ex As Exception

                            '            End Try

                            '        End If
                            '        If boolSTATSCV Then
                            '            Try
                            '                'do %CV: numSD / numMean
                            '                int2 = int2 + 1
                            '                If int3 < 3 Then
                            '                    var2 = "NA"
                            '                    boolHasNA = True 'for legend purposes
                            '                Else
                            '                    var1 = (numSD / numMean) * 100
                            '                    var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '                End If
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '                '.Selection.TypeText(var2)
                            '            Catch ex As Exception

                            '            End Try

                            '        End If
                            '        If boolSTATSBIAS and boolstatsmean  Then
                            '            'do %Bias: Mean/Nominal
                            '            Try
                            '                int2 = int2 + 1
                            '                If int3 < 3 Then
                            '                    var2 = "NA"
                            '                    boolHasNA = True 'for legend purposes
                            '                Else
                            '                    Try
                            '                        var1 = ((numMean / numNomConc) - 1) * 100
                            '                        var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '                    Catch ex As Exception
                            '                        var2 = "NA"
                            '                        boolHasNA = True 'for legend purposes
                            '                    End Try

                            '                End If
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '                '.Selection.TypeText(var2)
                            '            Catch ex As Exception

                            '            End Try

                            '        End If
                            '        If boolSTATSN Then
                            '            Try
                            '                'n
                            '                int2 = int2 + 1
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col4).Select()
                            '                'var2 = RoundToDecimal(var1, 0)
                            '                '.Selection.TypeText(CStr(int3))
                            '                '.Selection.TypeText(CStr(nI))

                            '            Catch ex As Exception

                            '            End Try
                            '        End If

                            '    End If
                            'End If

                            'column 5
                            'get average of peak areas for RS
                            strF = "NOMCONC = " & varNom
                            Erase rows2b
                            rows2b = tbl2b.Select(strF)
                            int2 = rows2b.Length

                            var1 = rows2b(0).Item("BOOLINTSTD")
                            If var1 = -1 Then
                                boolIS = True
                            Else
                                boolIS = False
                            End If
                            strF = ""
                            For Count4 = 0 To int2 - 1
                                var2 = rows2b(Count4).Item("ANALYTEINDEX")
                                var3 = rows2b(Count4).Item("MASTERASSAYID")
                                var4 = rows2b(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                                If Count4 <> int2 - 1 Then
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ") OR "
                                Else
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ")"
                                End If
                            Next
                            Erase rows3
                            rows3 = tbl1.Select(strF)
                            int3 = rows3.Length 'debug
                            nI = int3

                            strF = ""
                            For Count4 = 0 To int2 - 1
                                var2 = rows2b(Count4).Item("ANALYTEINDEX")
                                var3 = rows2b(Count4).Item("MASTERASSAYID")
                                var4 = rows2b(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                                If Count4 <> int2 - 1 Then
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND (ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL)) OR "
                                Else
                                    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND (ELIMINATEDFLAG = 'N' OR ELIMINATEDFLAG IS NULL))"
                                End If
                            Next
                            Erase rows3E
                            rows3E = tbl1.Select(strF)
                            nE = rows3E.Length 'debug

                            'get excluded

                            'For Count4 = 0 To intRowsT - 1
                            '    If Count4 > int3 - 1 Then
                            '        var1 = "NA"
                            '        boolHasNA = True 'for legend purposes
                            '    Else
                            '        If boolIS Then
                            '            var1 = rows3(Count4).Item("INTERNALSTANDARDAREA")
                            '        Else
                            '            var1 = rows3(Count4).Item("ANALYTEAREA")
                            '        End If
                            '    End If
                            '    '.Selection.Tables.Item(1).Cell(int1 + intInc + Count4, col5).Select()
                            '    '.Selection.TypeText(CStr(var1))
                            '    ''''''''''''''console.writeline(CStr(rows1(Count4).Item("RUNSAMPLESEQUENCENUMBER")) & ", " & var1)

                            '    If Count4 > int3 - 1 Then
                            '    Else
                            '        var2 = NZ(rows3(Count4).Item("ELIMINATEDFLAG"), "N")
                            '        'check for outlier
                            '        If StrComp(var2, "Y", vbTextCompare) = 0 Then
                            '            intLeg = intLeg + 1
                            '            strA = ChrW(intLeg + intLegStart)
                            '            str1 = "Value excluded from summary statistics because it is a statistical outlier according to the [OUTLIERMETHOD]."
                            '            'search for str1 in arrLegend
                            '            If intLeg = 1 Then
                            '                arrLegend(1, intLeg) = strA
                            '                arrLegend(2, intLeg) = str1
                            '                arrLegend(3, intLeg) = True
                            '                ctLegend = ctLegend + 1
                            '            Else
                            '                boolPro = True
                            '                For Count5 = 1 To intLeg - 1
                            '                    str2 = arrLegend(2, Count5)
                            '                    If StrComp(str1, str2, CompareMethod.Text) = 0 Then 'abort
                            '                        intLeg = intLeg - 1
                            '                        strA = arrLegend(1, Count5)
                            '                        boolPro = False
                            '                        Exit For
                            '                    End If
                            '                Next
                            '                If boolPro Then
                            '                    arrLegend(1, intLeg) = strA
                            '                    arrLegend(2, intLeg) = str1
                            '                    arrLegend(3, intLeg) = True
                            '                    ctLegend = ctLegend + 1
                            '                End If
                            '            End If
                            '            '.Selection.Font.Bold = True
                            '            '.Selection.Font.Color = Word.WdColor.wdColorRed
                            '            ''.Selection.TypeText(Text:=CStr(DisplayNum(arrBCStdActual(2, Count5), LSigFig)))
                            '            ''.Selection.TypeText(Text:=CStr(var2))
                            '            '.Selection.Font.Superscript = True
                            '            ''fontsize = .Selection.Font.Size
                            '            '.Selection.Font.Size = 12
                            '            ''.Selection.TypeText(strA)
                            '            '.Selection.TypeText(" " & strA)
                            '            '.Selection.Font.Superscript = False
                            '            '.Selection.Font.Size = fontsize
                            '            ''.Selection.TypeText Text:="NR"

                            '        End If

                            '    End If
                            'Next

                            int8 = -1
                            If boolQCREPORTACCVALUES Then
                            Else
                                If boolE(Count3) Then
                                    int8 = int8 + 1
                                End If
                            End If

                            intStart = int8

                            If boolSTATSMEAN Then
                                Try
                                    'do Mean

                                    If boolIS Then
                                        var1 = MeanDR(rows3E, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                                    Else
                                        var1 = MeanDR(rows3E, "ANALYTEAREA", False, "gaga", False, False)
                                    End If
                                    int8 = int8 + 1
                                    '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                                    var2 = RoundToDecimal(var1, 0)
                                    '.Selection.TypeText(Format(var2, "0"))
                                    numB = var2
                                    numMean = numB
                                Catch ex As Exception

                                End Try
                            End If
                            'If boolSTATSSD Then
                            '    Try
                            '        'do SD
                            '        int8 = int8 + 1
                            '        If int3 < 3 Then
                            '            var2 = "NA"
                            '            boolHasNA = True 'for legend purposes
                            '        Else
                            '            If boolIS Then
                            '                var1 = StdDevDR(rows3E, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                            '            Else
                            '                var1 = StdDevDR(rows3E, "ANALYTEAREA", False, "gaga", False, False)
                            '            End If
                            '            var2 = Format(RoundToDecimal(var1, 0), "0")
                            '            numSD = var2
                            '        End If
                            '        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '        '.Selection.TypeText(var2)
                            '    Catch ex As Exception

                            '    End Try
                            'End If
                            'If boolSTATSCV Then
                            '    Try
                            '        'do %CV: ((numSD/numMean)*100
                            '        int8 = int8 + 1
                            '        If int3 < 3 Then
                            '            var2 = "NA"
                            '            boolHasNA = True 'for legend purposes
                            '        Else
                            '            var1 = (numSD / numMean) * 100
                            '            var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '        End If
                            '        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '        '.Selection.TypeText(var2)
                            '    Catch ex As Exception

                            '    End Try
                            'End If
                            'If boolSTATSBIAS and boolstatsmean  Then
                            '    Try
                            '        'do %Bias: ((Mean/NomConc)*100
                            '        int8 = int8 + 1
                            '        If int3 < 3 Then
                            '            var2 = "NA"
                            '            boolHasNA = True 'for legend purposes
                            '        Else
                            '            var1 = ((numMean / numNomConc) - 1) * 100
                            '            var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '        End If
                            '        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '        '.Selection.TypeText(var2)
                            '    Catch ex As Exception

                            '    End Try
                            'End If
                            'If boolSTATSN Then
                            '    Try
                            '        'n
                            '        int8 = int8 + 1
                            '        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '        '.Selection.TypeText(CStr(nE))
                            '    Catch ex As Exception

                            '    End Try
                            'End If

                            'If boolQCREPORTACCVALUES Then
                            'Else
                            '    If boolE(Count3) Then
                            '        int8 = int8 + 2

                            '        If boolSTATSMEAN Then
                            '            Try
                            '                'do Mean

                            '                If boolIS Then
                            '                    var1 = MeanDR(rows3, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                            '                Else
                            '                    var1 = MeanDR(rows3, "ANALYTEAREA", False, "gaga", False, False)
                            '                End If
                            '                int8 = int8 + 1
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '                var2 = RoundToDecimal(var1, 0)
                            '                '.Selection.TypeText(Format(var2, "0"))
                            '                numBI = var2
                            '                numMean = numBI
                            '            Catch ex As Exception

                            '            End Try
                            '        End If
                            '        If boolSTATSSD Then
                            '            Try
                            '                'do SD
                            '                int8 = int8 + 1
                            '                If int3 < 3 Then
                            '                    var2 = "NA"
                            '                    boolHasNA = True 'for legend purposes
                            '                Else
                            '                    If boolIS Then
                            '                        var1 = StdDevDR(rows3, "INTERNALSTANDARDAREA", False, "gaga", False, True)
                            '                    Else
                            '                        var1 = StdDevDR(rows3, "ANALYTEAREA", False, "gaga", False, False)
                            '                    End If
                            '                    var2 = Format(RoundToDecimal(var1, 0), "0")
                            '                    numSD = var2
                            '                End If
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '                '.Selection.TypeText(var2)
                            '            Catch ex As Exception

                            '            End Try
                            '        End If
                            '        If boolSTATSCV Then
                            '            Try
                            '                'do %CV: ((numSD/numMean)*100
                            '                int8 = int8 + 1
                            '                If int3 < 3 Then
                            '                    var2 = "NA"
                            '                    boolHasNA = True 'for legend purposes
                            '                Else
                            '                    var1 = (numSD / numMean) * 100
                            '                    var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '                End If
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '                '.Selection.TypeText(var2)
                            '            Catch ex As Exception

                            '            End Try
                            '        End If
                            '        If boolSTATSBIAS and boolstatsmean  Then
                            '            Try
                            '                'do %Bias: ((Mean/NomConc)*100
                            '                int8 = int8 + 1
                            '                If int3 < 3 Then
                            '                    var2 = "NA"
                            '                    boolHasNA = True 'for legend purposes
                            '                Else
                            '                    var1 = ((numMean / numNomConc) - 1) * 100
                            '                    var2 = Format(RoundToDecimal(var1, 1), "0.0")
                            '                End If
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '                '.Selection.TypeText(var2)
                            '            Catch ex As Exception

                            '            End Try
                            '        End If
                            '        If boolSTATSN Then
                            '            Try
                            '                'n
                            '                int8 = int8 + 1
                            '                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int8, col5).Select()
                            '                '.Selection.TypeText(CStr(nI))
                            '            Catch ex As Exception

                            '            End Try
                            '        End If
                            '    End If
                            'End If

                            'column 6
                            int2 = 0
                            If boolQCREPORTACCVALUES Then
                                '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col6).Select()
                                '.Selection.Tables.Item(1).Cell(intRowCol6, col6).Select()
                                var1 = RoundToDecimal(numA / numB * 100, 1)
                                arr1(1, Count3 + 1) = var1
                                '.Selection.TypeText(CStr(Format(var1, "0.0")))
                            Else
                                If boolE(Count3) Then
                                    '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col6).Select()
                                    '.Selection.Tables.Item(1).Cell(intRowCol6, col6).Select()
                                    var1 = RoundToDecimal(numA / numB * 100, 1)
                                    arr1(1, Count3 + 1) = var1
                                    '.Selection.TypeText(CStr(Format(var1, "0.0")))

                                    '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col6).Select()
                                    '.Selection.Tables.Item(1).Cell(intRowCol6I, col6 + 1).Select()
                                    var1 = RoundToDecimal(numAI / numBI * 100, 1)
                                    arr1(2, Count3 + 1) = var1
                                    '.Selection.TypeText(CStr(Format(var1, "0.0")))
                                Else
                                    If boolETot Then
                                        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col6).Select()
                                        '.Selection.Tables.Item(1).Cell(intRowCol6, col6).Select()
                                        var1 = RoundToDecimal(numA / numB * 100, 1)
                                        arr1(1, Count3 + 1) = var1
                                        arr1(2, Count3 + 1) = var1
                                        '.Selection.TypeText(CStr(Format(var1, "0.0")))

                                        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col6).Select()
                                        '.Selection.Tables.Item(1).Cell(intRowCol6, col6 + 1).Select()
                                        var1 = RoundToDecimal(numA / numB * 100, 1)
                                        arr1(1, Count3 + 1) = var1
                                        arr1(2, Count3 + 1) = var1
                                        '.Selection.TypeText(CStr(Format(var1, "0.0")))
                                    Else
                                        '.Selection.Tables.Item(1).Cell(int1 + intInc + intRowsT + 1 + int2, col6).Select()
                                        '.Selection.Tables.Item(1).Cell(intRowCol6, col6).Select()
                                        var1 = RoundToDecimal(numA / numB * 100, 1)
                                        arr1(1, Count3 + 1) = var1
                                        arr1(2, Count3 + 1) = var1
                                        '.Selection.TypeText(CStr(Format(var1, "0.0")))
                                    End If

                                End If
                            End If

                            'intInc = intInc + intRowsT + 4 + 2
                            intInc = intInc + intRowsT + int8 + 1 + 2


                        Next

                        'now enter Mean/SD/%CV lables



                        int8 = 0
                        'fill arr2
                        ReDim arr2(intNumLevels)
                        For Count3 = 1 To intNumLevels
                            var1 = arr1(1, Count3)
                            arr2(Count3) = arr1(1, Count3)
                        Next

                        var1 = Mean(intNumLevels, arr2)
                        numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                        GetRecovery = CStr(numMean)

                        'If boolSTATSMEAN Then
                        '    Try
                        '        int8 = int8 + 1
                        '        'enter Mean
                        '        'numMean = Mean(intNumLevels, arr2)
                        '        .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6).Select()
                        '        var1 = Mean(intNumLevels, arr2)
                        '        numMean = SigFigOrDec(RoundToDecimal(var1, 5), LSigFig, False)


                        '        .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))

                        '    Catch ex As Exception

                        '    End Try
                        'End If
                        'If boolSTATSSD Then
                        '    Try
                        '        'enter SD
                        '        'numSD = StdDev(intNumLevels, arr2)
                        '        int8 = int8 + 1
                        '        .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6).Select()
                        '        var1 = StdDev(intNumLevels, arr2)
                        '        numSD = SigFigOrDec(RoundToDecimal(var1, 5), LSigFig, False)
                        '        .Selection.TypeText(CStr(DisplayNum(numSD, LSigFig, False)))

                        '    Catch ex As Exception

                        '    End Try
                        'End If
                        'If boolSTATSCV Then
                        '    Try
                        '        'enter %CV
                        '        int8 = int8 + 1
                        '        .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6).Select()
                        '        var1 = Format(RoundToDecimal(numSD / numMean * 100, 1), "0.0")
                        '        .Selection.TypeText(CStr(var1))

                        '    Catch ex As Exception

                        '    End Try
                        'End If
                        'If boolSTATSBIAS and boolstatsmean  Then
                        '    Try
                        '        'enter %Bias
                        '        int8 = int8 + 1
                        '        .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6).Select()
                        '        var1 = Format(RoundToDecimal(((numMean / numNomConc) - 1) * 100, 1), "0.0")
                        '        .Selection.TypeText(CStr(var1))

                        '    Catch ex As Exception

                        '    End Try
                        'End If
                        'If boolSTATSN Then
                        '    Try
                        '        'enter n
                        '        int8 = int8 + 1
                        '        .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6).Select()
                        '        var1 = intNumLevels
                        '        .Selection.TypeText(CStr(var1))

                        '    Catch ex As Exception

                        '    End Try
                        'End If

                        'If boolQCREPORTACCVALUES Then
                        'Else
                        '    If boolETot Then
                        '        int8 = 0
                        '        'fill arr2
                        '        ReDim arr2(intNumLevels)
                        '        For Count3 = 1 To intNumLevels
                        '            arr2(Count3) = arr1(2, Count3)
                        '        Next
                        '        If boolSTATSMEAN Then
                        '            Try
                        '                int8 = int8 + 1
                        '                'enter Mean
                        '                'numMean = Mean(intNumLevels, arr2)
                        '                .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6 + 1).Select()
                        '                var1 = Mean(intNumLevels, arr2)
                        '                numMean = SigFigOrDec(RoundToDecimal(var1, 5), LSigFig, False)
                        '                .Selection.TypeText(CStr(DisplayNum(numMean, LSigFig, False)))

                        '            Catch ex As Exception

                        '            End Try
                        '        End If
                        '        If boolSTATSSD Then
                        '            Try
                        '                'enter SD
                        '                'numSD = StdDev(intNumLevels, arr2)
                        '                int8 = int8 + 1
                        '                .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6 + 1).Select()
                        '                var1 = StdDev(intNumLevels, arr2)
                        '                numSD = SigFigOrDec(RoundToDecimal(var1, 5), LSigFig, False)
                        '                .Selection.TypeText(CStr(DisplayNum(numSD, LSigFig, False)))

                        '            Catch ex As Exception

                        '            End Try
                        '        End If
                        '        If boolSTATSCV Then
                        '            Try
                        '                'enter %CV
                        '                int8 = int8 + 1
                        '                .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6 + 1).Select()
                        '                var1 = Format(RoundToDecimal(numSD / numMean * 100, 1), "0.0")
                        '                .Selection.TypeText(CStr(var1))

                        '            Catch ex As Exception

                        '            End Try
                        '        End If
                        '        If boolSTATSBIAS and boolstatsmean  Then
                        '            Try
                        '                'enter %Bias
                        '                int8 = int8 + 1
                        '                .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6 + 1).Select()
                        '                var1 = Format(RoundToDecimal(((numMean / numNomConc) - 1) * 100, 1), "0.0")
                        '                .Selection.TypeText(CStr(var1))

                        '            Catch ex As Exception

                        '            End Try
                        '        End If
                        '        If boolSTATSN Then
                        '            Try
                        '                'enter n
                        '                int8 = int8 + 1
                        '                .Selection.Tables.Item(1).Cell(intStatsStart + int8, col6 + 1).Select()
                        '                var1 = intNumLevels
                        '                .Selection.TypeText(CStr(var1))

                        '            Catch ex As Exception

                        '            End Try
                        '        End If
                        '    End If
                        'End If


                        ''increase row position counter
                        'If Count2 = intNumRuns - 1 Then
                        '    int1 = int1 + intRowsX + int8 + 1  '4
                        'Else
                        '    int1 = int1 + intRowsX + int8 + 2 '5
                        'End If

                    Next

                End If
end1:
            Next
end2:

        Catch ex As Exception

        End Try



    End Function

    Function GetDilFactor(ByVal dtbl As System.Data.DataTable, ByVal idS As Int64, ByVal strAnal As String, ByVal idT As Int64)

        Dim intTableID As Short
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim tbl4 As System.Data.DataTable
        Dim strF As String
        Dim rows11() As DataRow
        Dim intRowsAnal As Short
        Dim intStart As Short
        Dim intEnd As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strDo As String
        Dim var1, var2, var3, var4
        Dim rowsX() As DataRow
        Dim strF2 As String
        Dim strS As String
        Dim rows2() As DataRow
        Dim dv2 As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim tblNumRuns As System.Data.DataTable
        Dim intNumRuns As Short
        Dim intNumLevels As Short
        Dim tblLevels As System.Data.DataTable

        GetDilFactor = "NA"

        tbl1 = tblAnalysisResultsHome
        tbl2 = tblAssignedSamples
        'tbl3 = tblAssignedSamplesHelper
        tbl4 = tblAnalytesHome

        strF = "IsIntStd = 'No'"
        rows11 = tblAnalytesHome.Select(strF)
        intRowsAnal = rows11.Length

        'find intStart and intEnd
        intStart = 0
        intEnd = 0
        For Count1 = 0 To tbl4.Rows.Count - 1
            str1 = tbl4.Rows(Count1).Item("AnalyteDescription")
            If StrComp(str1, strAnal, CompareMethod.Text) = 0 Then
                intStart = Count1 + 1
                intEnd = Count1 + 1
                Exit For
            End If
        Next


        For Count1 = intStart To intEnd 'intRowsAnal


            ''check if table is to be generated
            strDo = strAnal

            strF = "id_tblconfigreporttables = " & idT & " AND ID_TBLSTUDIES = " & idS & " AND  ID_TBLCONFIGREPORTTABLES = " & idT ' & " AND CHARANALYTE = '" & CleanText(strDo) & "'" ' AND ID_TBLREPORTTABLE = " & idTR
            rowsX = tbl2.Select(strF)
            If rowsX.Length = 0 Then
                Exit Function
            Else
            End If

            'setup tables
            var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
            var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
            var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
            strF2 = "ID_TBLSTUDIES = " & idS & " AND "
            strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & idT & " AND "
            'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
            strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
            strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "
            strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
            'for this function, need to sort by aliquot factor
            strS = "ALIQUOTFACTOR DESC"
            rows2 = tbl2.Select(strF2, strS)
            int1 = rows2.Length 'debug
            dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
            int1 = dv2.Count 'debug

            'find number of runs used
            tblNumRuns = dv2.ToTable("a", True, "RUNID")
            intNumRuns = tblNumRuns.Rows.Count

            'establish table of level numbers
            tblLevels = dv2.ToTable("b", True, "NOMCONC", "ALIQUOTFACTOR")
            intNumLevels = tblLevels.Rows.Count

            For Count2 = 0 To intNumLevels - 1
                var1 = NZ(tblLevels.Rows(Count2).Item("ALIQUOTFACTOR"), 1)
                var2 = NZ(tblLevels.Rows(Count2).Item("NOMCONC"), 1) 'debug
                'var2 = RoundToDecimalRAFZ(1 / var1, 0)
                var2 = GetDilnFactor(CDec(var1)) '20190220 LEE
                If Count2 = 0 Then
                    GetDilFactor = CStr(var2) & " fold"
                Else
                    GetDilFactor = GetDilFactor & ", " & CStr(var2) & " fold"
                End If

            Next


        Next

    End Function


    Sub FillInterIntraQCStats()

        '''''''''''console.writeline("Start: " & Now)

        'ignore if study is method validation

        Dim strF As String
        Dim str1 As String
        Dim tbl As System.Data.DataTable
        Dim dr() As DataRow
        Dim boolMethVal As Boolean
        tbl = tblConfigReportType
        strF = "ID_TBLSTUDIES = " & id_tblStudies ' & " AND CHARREPORTTYPE = 'Sample Analysis'"

        tbl = tblReports
        strF = "ID_TBLSTUDIES = " & id_tblStudies ' & " AND CHARREPORTTYPE = 'Sample Analysis'"
        dr = tbl.Select(strF)
        If dr.Length = 0 Then
            boolMethVal = False
        Else
            str1 = NZ(dr(0).Item("CHARREPORTTYPE"), "Sample Analysis")
            If InStr(1, str1, "Method", CompareMethod.Text) > 0 Then
                boolMethVal = True
            Else
                boolMethVal = False
                'do during AssessQCs instead
                Exit Sub
            End If
        End If
       
        '******

        Dim numNomConc As Decimal
        Dim numSD As Decimal
        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9, var10
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
        Dim arrFP(20) 'FlagPercent array
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

        Dim arrInterQCAcc(100)
        Dim arrInterQCPrec(100)
        Dim arrIntraQCAcc(100)
        Dim arrIntraQCPrec(100)
        Dim intCtInter As Short
        Dim intCtIntra As Short
        Dim strAnal As String
        Dim numMeanA As Decimal
        Dim numBiasA As Decimal
        Dim numSDA As Decimal

        Dim dtbl As System.Data.DataTable

        dtbl = tblInterQCSum


        If dtbl.Columns.Contains("Value") Then

            dtbl.Clear()

        Else

            Dim col11 As New DataColumn
            str1 = "Value"
            col11.ColumnName = str1
            col11.DataType = System.Type.GetType("System.Decimal")
            col11.DefaultValue = 0
            col11.AllowDBNull = True
            dtbl.Columns.Add(col11)

            Dim col22 As New DataColumn
            str1 = "Type"
            col22.ColumnName = str1
            col22.DataType = System.Type.GetType("System.String")
            col22.AllowDBNull = True
            dtbl.Columns.Add(col22)

            Dim col33 As New DataColumn
            str1 = "Analyte"
            col33.ColumnName = str1
            col33.DataType = System.Type.GetType("System.String")
            col33.AllowDBNull = True
            dtbl.Columns.Add(col33)

        End If

        boolJustTable = False



        'dvDo = frmH.dgvReportTableConfiguration.DataSource
        strTName = "Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision"
        'intDo = FindRowDVByCol(strTName, dvDo, "Table")

        intTableID = 11
        dvDo = frmH.dgvReportTableConfiguration.DataSource
        strF = "id_tblconfigreporttables = " & intTableID
        intDo = FindRowDVNumByCol(intTableID, dvDo, "id_tblconfigreporttables")

        ' ''Get table name
        ''var1 = dvDo(intDo).Item("Table")
        ''strTName = NZ(var1, "[NONE]")

        ''***
        'intDo = FindRowDVNumByCol(idTR, dvDo, "ID_TBLREPORTTABLE")
        ''intLeg = 0
        ''intLegStart = 96
        ''boolPro = False

        ''Get table name
        ''var1 = dvDo(intDo).Item("Table")
        'var1 = dvDo(intDo).Item("CHARHEADINGTEXT")
        'strTName = NZ(var1, "[NONE]")

        ''get Temperature info
        'var1 = dvDo(intDo).Item("CHARSTABILITYPERIOD")
        'strTempInfo = NZ(var1, "[NONE]")


        tbl1 = tblAnalysisResultsHome
        Dim dvTbl1 As System.Data.DataView = New DataView(tbl1)
        tbl2 = tblAssignedSamples
        tbl4 = tblAnalytesHome

        'ensure data has been entered
        strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies ' & " AND ID_TBLREPORTTABLE = " & idTR
        rowsX = tbl2.Select(strF)
        If rowsX.Length = 0 Then

            Exit Sub
        End If

        'If rowsX.Length = 0 Then
        '    strM = "Creating Summary of Interpolated QC Standard Concentrations Table ...."
        '    frmH.lblProgress.Text = strM
        '    frmH.Refresh()
        '    MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
        '    GoTo end2
        'End If

        strF = "IsIntStd = 'No'"
        rows11 = tblAnalytesHome.Select(strF)
        intRowsAnal = rows11.Length

        'build tblAnova
        Dim col1 As New DataColumn
        col1.ColumnName = "Group"
        col1.DataType = System.Type.GetType("System.Int16")
        tblAnova.Columns.Add(col1)
        Dim col2 As New DataColumn
        col2.ColumnName = "Conc"
        col2.DataType = System.Type.GetType("System.Decimal")
        tblAnova.Columns.Add(col2)
        Dim col3 As New DataColumn
        col3.ColumnName = "NomConc"
        col3.DataType = System.Type.GetType("System.Decimal")
        tblAnova.Columns.Add(col3)


        For Count1 = 1 To intRowsAnal

            Dim arrLegend(4, 20)

            intCtInter = 0
            intCtIntra = 0


            tblAnova.Clear()

            'check if table is to be generated
            'strDo = arrAnalytes(1, Count1) 'record column name
            strDo = rows11(Count1 - 1).Item("ANALYTEDESCRIPTION")
            strAnal = strDo
            bool = dvDo.Item(intDo).Item(strAnal) 'find boolean value of dvDo column

            If bool Then 'continue
                'ensure data has been entered
                'strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & stranal ' & "' AND ID_TBLREPORTTABLE = " & idTR
                'rowsX = tbl2.Select(strF)

                strF = "id_tblconfigreporttables = " & intTableID & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARANALYTE = '" & CleanText(strAnal) & "'" ' AND ID_TBLREPORTTABLE = " & idTR
                rowsX = tbl2.Select(strF)
                If rowsX.Length = 0 Then
                    'strM = "Creating Summary of Interpolated QC Standard Concentrations Table ...."
                    'frmH.lblProgress.Text = strM
                    'frmH.Refresh()
                    'MsgBox("Samples have not been assigned to this table.", MsgBoxStyle.Information, "Samples have not been assigned...")
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

                'get strConcUnits
                int1 = FindRowDV("ULOQ Units", frmH.dgvWatsonAnalRef.DataSource)
                strConcUnits = NZ(frmH.dgvWatsonAnalRef(Count1, int1).Value, "ng/mL")

                int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

                If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    strConcUnits = str1
                End If

                'setup tables
                'legend
                'tbl1 = tblAnalysisResultsHome
                'Dim dvTbl1 As System.Data.DataView = New DataView(tbl1)
                'tbl2 = tblAssignedSamples
                'tbl4 = tblAnalytesHome
                var1 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteIndex")
                var2 = tbl4.Rows.Item(Count1 - 1).Item("MasterAssayID")
                var3 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteDescription")
                var4 = tbl4.Rows.Item(Count1 - 1).Item("AnalyteID")
                var5 = tbl4.Rows.Item(Count1 - 1).Item("intGroup")
                strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND "
                strF2 = strF2 & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                'strF2 = strF2 & "ID_TBLREPORTTABLE = " & idTR & " AND "
                'strF2 = strF2 & "ANALYTEINDEX = " & var1 & " AND "
                'strF2 = strF2 & "MASTERASSAYID = " & var2 ' & " AND "
                'strF2 = strF2 & "CHARANALYTE = '" & CleanText(cstr(var3)) & "' AND "
                strF2 = strF2 & "INTGROUP = " & var5 ' & " AND "
                'strF2 = strF2 & "BOOLINTSTD = 0"
                strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
                rows2 = tbl2.Select(strF2, strS)
                int1 = rows2.Length 'debug
                dv2 = New DataView(tbl2, strF2, strS, DataViewRowState.CurrentRows)
                int1 = dv2.Count 'debug

                'find number of runs used
                tblNumRuns = dv2.ToTable("a", True, "RUNID")
                intNumRuns = tblNumRuns.Rows.Count

                'establish table of level numbers
                tblLevels = dv2.ToTable("b", True, "NOMCONC", "CHARHELPER1")
                intNumLevels = tblLevels.Rows.Count

                'find number of table rows to generate
                intRowsX = 0
                For Count2 = 0 To intNumRuns - 1
                    '.Selection.Tables.item(1).Cell(int1, 1).Select()
                    'enter runid
                    var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")
                    '.Selection.TypeText(CStr(var10))
                    '.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCell)

                    'intRowsX = 0
                    For Count3 = 0 To intNumLevels - 1
                        varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                        var1 = tblLevels.Rows.Item(Count3).Item("CHARHELPER1")
                        If IsDBNull(varNom) Then
                            str1 = "Warning: Nominal Concentration for " & strDo & " for Assignment '" & var1 & "' in table:" & ChrW(10) & strTName & ChrW(10) & "has not been assigned."
                            str1 = str1 & ChrW(10) & "Nominal Concentration will be set to 1"
                            MsgBox(str1, MsgBoxStyle.Information, "Sample assignment incomplete...")
                            tblLevels.Rows.Item(Count3).BeginEdit()
                            tblLevels.Rows.Item(Count3).Item("NOMCONC") = 1
                            tblLevels.Rows.Item(Count3).EndEdit()
                            varNom = 1
                        End If
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
                Next

                'begin entering data'
                int1 = 5 'row position counter
                intLeg = 0
                ctQCLegend = 0
                ctDilLeg = 0
                ctLegend = 0
                strA = ""
                strB = ""
                For Count2 = 0 To intNumRuns - 1

                    'enter runid
                    var10 = tblNumRuns.Rows.Item(Count2).Item("RUNID")

                    'start filling in data by columns
                    'intRowsX = 0
                    For Count3 = 0 To intNumLevels - 1

                        intN = 0

                        varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")

                        ''determine hi and lo (nom*flagpercent)
                        'strF = "CONCENTRATION = '" & varNom & "'"
                        'rows10 = tblBCQCs.Select(strF)
                        'var1 = NZ(rows10(0).Item("FLAGPERCENT"), 15)
                        'var1 = CDec(var1)
                        'arrFP(Count3) = var1
                        'hi = SigFigOrDec(varNom + (varNom * var1 / 100), LSigFig, False)
                        'lo = SigFigOrDec(varNom - (varNom * var1 / 100), LSigFig, False)

                        'start entering data
                        dv2.RowFilter = ""
                        'don't know why, but must make a long filter here or
                        'both analytes get returned in dv2.rowfilter
                        'if study is method validation, then filter for id_tblConfigReportTables = 11
                        'otherwise id_tblConfigReportTables = 4
                        If boolMethVal Then
                            strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ID_TBLCONFIGREPORTTABLES = 11"
                        Else
                            strF = strF2 & " AND RUNID = " & var10 & " AND NOMCONC = " & varNom & " AND ID_TBLCONFIGREPORTTABLES = 4"
                        End If

                        dv2.RowFilter = strF
                        int2 = dv2.Count
                        If int2 = 0 Then
                        Else
                            'create rows1 from tbl1 which will contain data
                            strF = ""
                            'legend
                            'dv2 = 'tbl2 = tblAssignedSamples
                            'legend
                            'tbl1 = tblAnalysisResultsHome
                            'Dim dvTbl1 As System.Data.DataView = New DataView(tbl1)
                            'tbl2 = tblAssignedSamples
                            'tbl4 = tblAnalytesHome
                            For Count4 = 0 To dv2.Count - 1
                                var2 = dv2(Count4).Item("ANALYTEINDEX")
                                var3 = dv2(Count4).Item("MASTERASSAYID")
                                var4 = dv2(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                                Try
                                    var5 = dv2(Count4).Item("INTGROUP")
                                Catch ex As Exception
                                    var5 = ex.Message
                                End Try
                                var6 = NZ(dv2(Count4).Item("ANALYTEID"), -98765)

                                'If Count4 <> dv2.Count - 1 Then
                                '    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ") OR "
                                'Else
                                '    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ")"
                                'End If
                                Try
                                    If Count4 <> dv2.Count - 1 Then
                                        strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var6 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ") OR "
                                    Else
                                        strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var6 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & ")"
                                    End If
                                Catch ex As Exception
                                    strF = ex.Message
                                End Try
                            
                            Next
                            'Erase rows1
                            'rows1 = tbl1.Select(strF)
                            'int3 = rows1.Length

                            'legend
                            'tbl1 = tblAnalysisResultsHome
                            'Dim dvTbl1 As System.Data.DataView = New DataView(tbl1)
                            'tbl2 = tblAssignedSamples
                            'tbl4 = tblAnalytesHome

                            'now do rows actual
                            ''debug
                            'Console.WriteLine("Start")
                            'For Count4 = 0 To tbl1.Columns.Count - 1
                            '    Console.WriteLine(tbl1.Columns(Count4).ColumnName)
                            'Next
                            'Console.WriteLine("End")
                            strFActual = "(" & strF & ") AND ELIMINATEDFLAG = 'N'"
                            rowsActual = tbl1.Select(strFActual)
                            intN = rowsActual.Length

                            int8 = 0

                            intCtIntra = intCtIntra + 1

                            If boolSTATSMEAN Then
                                Try
                                    'enter Mean
                                    int8 = int8 + 1
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    'var1 = MeanDR(rows1, "CONCENTRATION", True, "ALIQUOTFACTOR", True)
                                    var1 = MeanDR(rowsActual, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                    numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    '.Selection.TypeText(CStr(numMean))


                                Catch ex As Exception

                                End Try
                            End If
                            If boolSTATSSD Then
                                Try
                                    'enter SD
                                    int8 = int8 + 1
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    'var1 = MeanDR(rows1, "CONCENTRATION", True, "ALIQUOTFACTOR", True)
                                    'var1 = StdDevDR(rows1, "CONCENTRATION", True, "ALIQUOTFACTOR", True)
                                    var1 = StdDevDR(rowsActual, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                    numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                    '.Selection.TypeText(CStr(numSD))

                                Catch ex As Exception

                                End Try
                            End If
                            If boolSTATSCV Then
                                Try
                                    'enter %CV
                                    int8 = int8 + 1
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    'var1 = Format(RoundToDecimal(((var2 / varNom) - 1) * 100, 1), "0.0")
                                    If numMean = 0 Then
                                        var1 = 0
                                    Else
                                        var1 = RoundToDecimal(((numSD / numMean) * 100), 3)
                                    End If
                                    var1 = Format(var1, "0.0")

                                    arrIntraQCPrec(intCtIntra) = var1

                                    '.Selection.TypeText(CStr(var1))
                                Catch ex As Exception

                                End Try
                            End If
                            If boolSTATSBIAS And boolSTATSMEAN Then
                                Try
                                    'enter %Bias
                                    int8 = int8 + 1
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    'var1 = Format(RoundToDecimal(((var2 / varNom) - 1) * 100, 1), "0.0")
                                    var1 = RoundToDecimal((((numMean / varNom) - 1) * 100), 3)
                                    var1 = Format(var1, "0.0")

                                    arrIntraQCAcc(intCtIntra) = var1

                                    '.Selection.TypeText(CStr(var1))
                                Catch ex As Exception

                                End Try
                            End If
                            If boolSTATSN Then
                                Try
                                    'enter n
                                    int8 = int8 + 1
                                    '.Selection.Tables.Item(1).Cell(int1 + intRowsX + int8, Count3 + 2).Select()
                                    '.Selection.TypeText(CStr(int2))
                                    '.Selection.TypeText(CStr(intN))
                                Catch ex As Exception

                                End Try
                            End If

                        End If

                    Next

                    ''increase row position counter
                    'If Count2 = intNumRuns - 1 Then
                    '    int1 = int1 + intRowsX + int8 + 1 '4
                    'Else
                    '    int1 = int1 + intRowsX + int8 + 2 '5
                    'End If

                    ''''wdd.visible = True

                Next

                '''''''''''console.writeline("End Intra: " & Now)

                Dim boolGo As Boolean

                var1 = 1 'debugging

                'begin evaluating for interrun data
                For Count3 = 0 To intNumLevels - 1
                    varNom = tblLevels.Rows.Item(Count3).Item("NOMCONC")
                    'strF = strF2 & "AND NOMCONC = " & varNom
                    'dv2.RowFilter = ""
                    'dv2.RowFilter = strF

                    'legend
                    'tbl1 = tblAnalysisResultsHome
                    'Dim dvTbl1 As System.Data.DataView = New DataView(tbl1)
                    'tbl2 = tblAssignedSamples
                    'tbl4 = tblAnalytesHome

                    strF = strF2 & "AND NOMCONC = " & varNom
                    dv2.RowFilter = ""
                    dv2.RowFilter = strF
                    int2 = dv2.Count
                    'create rows1 from tbl1 which will contain data
                    strF = ""
                    For Count4 = 0 To dv2.Count - 1
                        var2 = dv2(Count4).Item("ANALYTEINDEX")
                        var3 = dv2(Count4).Item("MASTERASSAYID")
                        var4 = dv2(Count4).Item("RUNSAMPLESEQUENCENUMBER")
                        var10 = dv2(Count4).Item("RUNID")
                        Try
                            var5 = dv2(Count4).Item("INTGROUP")
                        Catch ex As Exception
                            var5 = ex.Message
                        End Try
                        var6 = NZ(dv2(Count4).Item("ANALYTEID"), -98765)
                        'If Count4 <> dv2.Count - 1 Then
                        '    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND ELIMINATEDFLAG = 'N') OR "
                        'Else
                        '    strF = strF & "(RUNID = " & var10 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND ELIMINATEDFLAG = 'N')"
                        'End If

                        'can't use intGroup here because tbl1 does not have intGroup
                        Try
                            If Count4 <> dv2.Count - 1 Then
                                strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var6 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND ELIMINATEDFLAG = 'N') OR "
                            Else
                                strF = strF & "(RUNID = " & var10 & " AND ANALYTEID = " & var6 & " AND RUNSAMPLESEQUENCENUMBER = " & var4 & " AND ELIMINATEDFLAG = 'N')"
                            End If
                        Catch ex As Exception
                            strF = ex.Message
                        End Try

                    Next

                    Erase rows1
                    rows1 = tbl1.Select(strF)
                    int3 = rows1.Length

                    ''start ANOVA section
                    'dvAn = tblAnova.DefaultView

                    ''retrieve anova
                    'dvAn.RowFilter = ""
                    'dvAn.RowFilter = "NomConc = " & varNom
                    'tblAnGo.Clear()
                    'tblAnGo = dvAn.ToTable
                    ''ReturnAnova = ANOVA_OneWay(tblAnGo)

                    int2 = 4
                    'For Count4 = 5 To 13
                    Dim intCC As Short
                    intCC = intRow + 1

                    intCtInter = intCtInter + 1

                    For Count4 = intRow + 1 To intRow + 1 + 14 '8
                        int2 = int2 + 1
                        str1 = ""
                        boolGo = True
                        Select Case int2
                            Case 6 'enter Mean Obs Conc
                                str1 = "Mean Observed Conc."
                                var1 = MeanDR(rows1, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                numMean = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                var1 = numMeanA

                            Case 7 'Inter-Run S.D.
                                var1 = StdDevDR(rows1, "CONCENTRATION", True, "ALIQUOTFACTOR", True, False)
                                numSD = SigFigOrDecString(RoundToDecimalA(var1, LSigFig), LSigFig, False)
                                var1 = numSDA

                            Case 8 'Inter-Run %CV
                                If numMean = 0 Then
                                    var1 = 0
                                Else
                                    var1 = RoundToDecimal(((numSD / numMean) * 100), 3)
                                End If
                                var1 = Format(var1, "0.0")

                                arrInterQCPrec(intCtInter) = CDec(var1)

                            Case 9 'Inter-Run Bias
                                var1 = RoundToDecimal((((numMean / varNom) - 1) * 100), 3)
                                numBias = RoundToDecimal(var1, 1)

                                arrInterQCAcc(intCtInter) = numBias

                                var1 = numBiasA

                            Case 10 'n
                                If boolSTATSN Then
                                    str1 = CStr(int3)
                                Else
                                    boolGo = False
                                End If

                            Case 12 'Between Run Precision (%CV)
                                'str1 = CStr(ReturnAnova(0))
                            Case 13 'Within Run Precision (%CV)
                                'str1 = CStr(ReturnAnova(1))
                            Case 14 'Number of runs
                                'find number of runs by doing a distinct
                                tblZ = dv2.ToTable("a", True, "RUNID")
                                int1 = tblZ.Rows.Count
                                str1 = CStr(int1)
                        End Select
                        If boolGo Then
                            intCC = intCC + 1
                        End If

                    Next
                Next

            End If

end1:

            'fill dtbl
            For Count2 = 1 To intCtInter
                var1 = arrInterQCAcc(Count2)
                var1 = NZ(arrInterQCAcc(Count2), "")
                If Len(var1) = 0 Then
                Else
                    Dim nrow As DataRow = dtbl.NewRow
                    nrow.BeginEdit()
                    nrow("Value") = var1
                    nrow("Type") = "InterAcc"
                    nrow("Analyte") = strAnal
                    nrow.EndEdit()
                    dtbl.Rows.Add(nrow)
                End If
            Next
            For Count2 = 1 To intCtInter
                var1 = arrInterQCPrec(Count2)
                var1 = NZ(arrInterQCPrec(Count2), "")
                If Len(var1) = 0 Then
                Else
                    Dim nrow As DataRow = dtbl.NewRow
                    nrow.BeginEdit()
                    nrow("Value") = var1
                    nrow("Type") = "InterPrec"
                    nrow("Analyte") = strAnal
                    nrow.EndEdit()
                    dtbl.Rows.Add(nrow)
                End If
            Next

            For Count2 = 1 To intCtIntra

                var1 = NZ(arrIntraQCAcc(Count2), "")
                If Len(var1) = 0 Then
                Else
                    Dim nrow As DataRow = dtbl.NewRow
                    nrow.BeginEdit()
                    nrow("Value") = var1
                    nrow("Type") = "IntraAcc"
                    nrow("Analyte") = strAnal
                    nrow.EndEdit()
                    dtbl.Rows.Add(nrow)
                End If

            Next
            For Count2 = 1 To intCtIntra
                var1 = arrIntraQCPrec(Count2)
                var1 = NZ(arrIntraQCPrec(Count2), "")
                If Len(var1) = 0 Then
                Else
                    Dim nrow As DataRow = dtbl.NewRow
                    nrow.BeginEdit()
                    nrow("Value") = var1
                    nrow("Type") = "IntraPrec"
                    nrow("Analyte") = strAnal
                    nrow.EndEdit()
                    dtbl.Rows.Add(nrow)
                End If
            Next

            '''''''''''console.writeline("End Inter: " & Now)

        Next


end2:

        '''''''''''console.writeline("End Inter: " & Now)


    End Sub

    Function GetIntraAccMin(ByVal strAnal As String) As String

        GetIntraAccMin = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'IntraAcc' AND Analyte = '" & strAnal & "'"
        strS = "Value ASC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetIntraAccMin = CStr(rows(0).Item("Value"))
        End If

    End Function

    Function GetIntraAccMax(ByVal strAnal As String) As String

        GetIntraAccMax = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'IntraAcc' AND Analyte = '" & strAnal & "'"
        strS = "Value DESC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetIntraAccMax = CStr(rows(0).Item("Value"))
        End If

    End Function

    Function GetInterAccMin(ByVal strAnal As String) As String

        GetInterAccMin = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'InterAcc' AND Analyte = '" & strAnal & "'"
        strS = "Value ASC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetInterAccMin = CStr(rows(0).Item("Value"))
        End If

    End Function

    Function GetInterAccMax(ByVal strAnal As String) As String

        GetInterAccMax = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'InterAcc' AND Analyte = '" & strAnal & "'"
        strS = "Value DESC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetInterAccMax = CStr(rows(0).Item("Value"))
        End If

    End Function


    '****

    Function GetIntraPrecMin(ByVal strAnal As String) As String

        GetIntraPrecMin = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'IntraPrec' AND Analyte = '" & strAnal & "'"
        strS = "Value ASC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetIntraPrecMin = CStr(rows(0).Item("Value"))
        End If

    End Function

    Function GetIntraPrecMax(ByVal strAnal As String) As String

        GetIntraPrecMax = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'IntraPrec' AND Analyte = '" & strAnal & "'"
        strS = "Value DESC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetIntraPrecMax = CStr(rows(0).Item("Value"))
        End If

    End Function

    Function GetInterPrecMin(ByVal strAnal As String) As String

        GetInterPrecMin = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'InterPrec' AND Analyte = '" & strAnal & "'"
        strS = "Value ASC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetInterPrecMin = CStr(rows(0).Item("Value"))
        End If

    End Function

    Function GetInterPrecMax(ByVal strAnal As String) As String

        GetInterPrecMax = "NA"

        Dim Count1 As Short
        Dim strF As String
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow

        dtbl = tblInterQCSum
        If dtbl.Rows.Count = 0 Then
            Exit Function
        End If

        strF = "Type = 'InterPrec' AND Analyte = '" & strAnal & "'"
        strS = "Value DESC"
        rows = dtbl.Select(strF, strS)
        If rows.Length = 0 Then
            Exit Function
        Else
            GetInterPrecMax = CStr(rows(0).Item("Value"))
        End If

    End Function

    '****

    Function GetCalibrStds(ByVal con As ADODB.Connection, ByVal arrA As Array, ByVal intPos As Short, ByVal idR As Short, ByVal idS2 As Int64) As String

        GetCalibrStds = ""

        'get this information from tblBCStds
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intRows As Short
        Dim MasterAssayID As Long
        Dim AnalyteIndex As Long
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim var1
        Dim num1 As Single
        Dim strLL As String

        MasterAssayID = arrA(12, intPos)
        AnalyteIndex = arrA(3, intPos)

        tbl1 = tblBCStds

        'strF = "STUDYID = " & id_tblStudies & " AND MASTERASSAYID = " & MasterAssayID & " AND ANALYTEINDEX = " & AnalyteIndex
        strF = "MASTERASSAYID = " & MasterAssayID & " AND ANALYTEINDEX = " & AnalyteIndex

        strS = "CONCENTRATION ASC"

        Dim dv1 As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
        Dim tbl2 As System.Data.DataTable = dv1.ToTable("a", True, "CONCENTRATION")
        intRows = tbl2.Rows.Count

        For Count1 = 0 To intRows - 1
            var1 = tbl2.Rows(Count1).Item("CONCENTRATION")
            If IsNumeric(var1) Then
                num1 = CSng(var1)
                If num1 < 100 Then
                    strLL = CStr(SigFigOrDecString(num1, LSigFig, False))
                Else
                    strLL = var1
                End If
            Else
                strLL = var1
            End If

            If Count1 = intRows - 1 Then
                GetCalibrStds = GetCalibrStds & " and " & strLL
            Else
                GetCalibrStds = GetCalibrStds & strLL & ", "
            End If
        Next

        'get units
        Dim dv As System.Data.DataView
        Dim strUnits As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim str1 As String

        dgv = frmH.dgvWatsonAnalRef

        dv = dgv.DataSource
        strUnits = "NA"
        strF = "LLOQ Units"
        int1 = FindRowDVByCol(strF, dv, "Item")
        If int1 = -1 Then
        Else
            strUnits = NZ(dgv(intPos, int1).Value, "NA")
        End If

        int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
        str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

        If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
        Else
            strUnits = str1
        End If

        GetCalibrStds = GetCalibrStds & " " & strUnits


    End Function

    Function GetQCs(ByVal con As ADODB.Connection, ByVal arrA As Array, ByVal intPos As Short, ByVal idR As Short, ByVal idS2 As Int64) As String

        GetQCs = ""

        'get this information from Assigned Samples table
        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim dtbl3 As System.Data.DataTable
        Dim dtblD As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim rows3() As DataRow
        Dim rowsD() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim strFD As String
        Dim intRows As Short
        Dim intCols As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolGo As Boolean
        Dim int1 As Short
        Dim strV As String
        Dim var1, var2, var3
        Dim strS As String
        Dim num1 As Single
        Dim strLL As String

        Dim dgv As DataGridView

        dgv = frmH.dgvMethodValData
        intRows = dgv.Rows.Count

        dtbl1 = tblAssignedSamples
        dtbl2 = tblAssignedSamplesHelper
        dtbl3 = tblConfigReportTables
        dtblD = tblMethodValidationData

        '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=Original AnalyteDescription

        'get lloq
        strF1 = "ID_TBLSTUDIES = " & idS2 & " AND ID_TBLCONFIGREPORTTABLES = " & idR & " AND ANALYTEINDEX = " & arrA(3, intPos) & " AND MASTERASSAYID = " & arrA(12, intPos) ' & " AND CHARHELPER1 = 'QC LLOQ'"
        strS = "NOMCONC ASC"
        'rows1 = dtbl1.Select(strF1)

        'now make unique dv
        Dim dv1 As System.Data.DataView = New DataView(dtbl1, strF1, strS, DataViewRowState.CurrentRows)
        Dim tblU As System.Data.DataTable = dv1.ToTable("a", True, "CHARHELPER1", "NOMCONC")
        intRows = tblU.Rows.Count
        For Count1 = 0 To intRows - 1
            var1 = NZ(tblU.Rows(Count1).Item("NOMCONC"), 1)
            If IsNumeric(var1) Then
                num1 = CSng(var1)
                If num1 < 100 Then
                    strLL = CStr(SigFigOrDecString(num1, LSigFig, False))
                Else
                    strLL = var1
                End If
            Else
                strLL = var1
            End If

            If Count1 = intRows - 1 Then
                GetQCs = GetQCs & " and " & strLL
            Else
                GetQCs = GetQCs & strLL & ", "
            End If

        Next

        'get units
        Dim dv As System.Data.DataView
        Dim strF As String
        Dim strUnits As String
        dgv = frmH.dgvWatsonAnalRef

        dv = dgv.DataSource
        strUnits = "NA"
        strF = "LLOQ Units"
        int1 = FindRowDVByCol(strF, dv, "Item")
        If int1 = -1 Then
        Else
            strUnits = NZ(dgv(intPos, int1).Value, "NA")
        End If

        int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
        str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

        If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
        Else
            strUnits = str1
        End If

        GetQCs = GetQCs & " " & strUnits



    End Function

    Function GetLLOQ(ByVal intPos As Short, ByVal strAnal As String) As String

        GetLLOQ = ""

        Dim int1 As Short
        Dim var1
        Dim strF As String
        Dim strLL As String
        Dim strUnits As String
        Dim ct1 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView

        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource

        strF = "LLOQ"
        strLL = "NA"
        strUnits = "NA"

        Dim num1 As Single
        Dim str1 As String

        int1 = FindRowDVByCol(strF, dv, "Item")

        '20910206: LEE
        If intPos > frmH.dgvWatsonAnalRef.ColumnCount - 1 Then
            GetLLOQ = "NA"
        Else            var1 = NZ(dgv(intPos, int1).Value, "NA")
            If IsNumeric(var1) Then
                num1 = CSng(var1)
                If num1 < 100 Then
                    strLL = CStr(SigFigOrDecString(num1, LSigFig, False))
                Else
                    strLL = var1
                End If
            Else
                strLL = var1
            End If

            strF = "LLOQ Units"
            int1 = FindRowDVByCol(strF, dv, "Item")
            strUnits = NZ(dgv(intPos, int1).Value, "NA")

            int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
            str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

            If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            Else
                strUnits = str1
            End If

            GetLLOQ = strLL & " " & strUnits

        End If

    End Function

    Function GetULOQ(ByVal intPos As Short, ByVal strAnal As String) As String

        GetULOQ = ""

        Dim int1 As Short
        Dim var1
        Dim strF As String
        Dim strLL As String
        Dim strUnits As String
        Dim ct1 As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim num1 As Single
        Dim str1 As String

        strLL = "NA"
        strUnits = "NA"

        dgv = frmH.dgvWatsonAnalRef
        dv = dgv.DataSource

        strF = "ULOQ"
        int1 = FindRowDVByCol(strF, dv, "Item")
        If intPos > dgv.ColumnCount - 1 Then
            GetULOQ = ""
        Else
            var1 = NZ(dgv(intPos, int1).Value, "NA")
            If IsNumeric(var1) Then
                num1 = CSng(var1)
                If num1 < 100 Then
                    strLL = CStr(SigFigOrDecString(num1, LSigFig, False))
                Else
                    strLL = var1
                End If
            Else
                strLL = var1
            End If
            strF = "ULOQ Units"
            int1 = FindRowDVByCol(strF, dv, "Item")
            strUnits = NZ(dgv(intPos, int1).Value, "NA")

            int1 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
            str1 = NZ(frmH.dgvStudyConfig(1, int1).Value, "")

            If Len(str1) = 0 Or StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            Else
                strUnits = str1
            End If

            GetULOQ = strLL & " " & strUnits
        End If
       

    End Function

End Module
