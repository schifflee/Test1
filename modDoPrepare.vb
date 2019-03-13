Option Compare Text

'slight change for git clone testing c

Module modDoPrepare

    Function DoPrepare(ByVal cn As ADODB.Connection) As Boolean


        'daDoPr

        Dim BACStudy As String
        'Dim 'frm As New 'frmprogress_01
        'Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim rs1 As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        Dim rs3 As New ADODB.Recordset
        Dim rsRS As New ADODB.Recordset
        Dim rsFindNomConc As New ADODB.Recordset
        Dim dbPath As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim wAnalyteID As Long
        'Dim wWStudyName As String
        'Dim ctAnalytes As short
        'Dim ctAnalytes_IS As short
        Dim ctAnalyticalRuns As Short

        Dim strF As String
        Dim strS As String

        Dim dtNow As Date = Now

        'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET

        Dim arrAnalyticalRuns(14, 500)
        '1=RUNID, 2=NOTEBOOK-PAGENUMBER, 3=EXTRACTIONDATE, 4=RUNSTARTDATE, 5=ANAREGSTATUSDESC
        '6=RUNDESCRIPTION, 7=ACCEPTREJECTREASON, 8=ANALYTE, 9=RUNTYPEID, 10=NM, 11=VEC, 12=ANALYTEID, 13=BOOLINTHISRUNSASSAYID, 14=RUNANALYTEREGRESSIONSTATUS

        Dim strA As String

        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim Count6 As Short
        Dim Count10 As Short
        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int5 As Short
        Dim int6 As Short
        Dim int10 As Short
        Dim int20 As Short
        Dim int30 As Short
        Dim int40 As Short
        Dim int50 As Short
        Dim int60 As Short
        Dim arrRegCon()
        Dim arrTemp(2, 50)
        Dim num1 As Decimal
        Dim num2 As Object
        Dim num3 As Object
        Dim arrBCStdActual()
        Dim arrLegend(1, 10)
        Dim ctLegend As Short
        Dim lng1 As Long
        Dim lng2 As Long
        Dim strPwd As String
        Dim fd, fi
        Dim boolGo As Boolean
        Dim strName As String
        Dim ctRows As Short
        Dim dtbl As System.Data.DataTable
        Dim boolRO As Boolean
        Dim dg As DataGrid
        Dim inttemprows As Short
        Dim numSum As Object
        Dim numMean As Object
        Dim numSD As Object
        Dim dv As System.Data.DataView
        Dim fld As ADODB.Field
        Dim drows() As DataRow
        'Dim numQCLevels As Short
        'Dim numRepDilnQC As Short
        Dim numRepQC As Short
        Dim maxRep As Short
        Dim boolI As Boolean
        Dim str_cbxStudy As String
        Dim boolD As Boolean
        Dim rsC As New ADODB.Recordset
        Dim intRows As Short

        Dim rsBCStds As New ADODB.Recordset
        Dim rs4 As New ADODB.Recordset
        Dim rs20 As New ADODB.Recordset

        Dim intAR, intARAnalytes As Short
        Dim intGroup As Short
        'Dim intUBAA As Short = 18 'upper bound of first paramter of arrAnalytes

        'public variable
        intUBAA = 20 '18 'upper bound of first paramter of arrAnalytes

        Dim vS, vM, vH, vD
        Dim sDate As Int64

        ID_QATEMPID = 0

        DoPrepare = True


        'GoTo skipLarry1

        '20150911: Larry
        'new Group paradigm to replace MasterAssayID paradigm
        'the first thing to do is establish tblCalStdGroupsAcc and tblQCStdGroups

        'tblCalStdGroupsAcc	    tblCalStdGroupAssayIDs	tblAnalyteGroups	    tblQCStdGroups	    tblQCStdGroupAssayIDs
        'ANALYTEDESCRIPTION	    ANALYTEDESCRIPTION	    ANALYTEDESCRIPTION	    ANALYTEDESCRIPTION	ANALYTEDESCRIPTION
        'ANALYTEDESCRIPTION_C	INTSTD	                INTSTD	                INTSTD	            INTSTD
        'INTSTD	                ANALYTEID	            GROUP	                ANALYTEID	        ANALYTEID
        'ANALYTEID	            ANALYTEINDEX	        ANALYTEDESCRIPTION_C	ANALYTEINDEX	    ANALYTEINDEX
        'ANALYTEINDEX	        ASSAYID		            ASSAYID	                ASSAYID
        'ASSAYID	            MASTERASSAYID		    MASTERASSAYID	        MASTERASSAYID
        'MASTERASSAYID	        RUNID		            LEVELNUMBER	            RUNID
        'LEVELNUMBER	        RUNDATE		            CONCENTRATION	        RUNDATE
        'CONCENTRATION	        GROUP		            RUNID	                GROUP
        'RUNID	                RUNTYPEID		        RUNDATE	                RUNTYPEID
        'RUNDATE                GROUP
        'GROUP()

        'ALSO (NDL 3-Feb-2016)
        'tblAnalyteIDs - all AssayID's in current study
        'tblMatrices   - all Matrices in current study

        'tblAnalyteIDs          tblMatrices
        'ANALYEID               MATRIX

        'to enable group paradigm, enter true
        boolUseGroups = True

        'first get watson version info

        'DBTABLEVER: 7.4
        'DBQUERYVER: 5.4
        'JETVER: 
        'VBVER: 
        'KEYFIELD: 
        'JETBITS: 
        'VBBITS: 
        'ENVIRONMENT: WATP
        'WATSONVER: 7.4.1

        If boolAccess Then
            str1 = "SELECT VERSIONCONTROL.* "
            str2 = "FROM VERSIONCONTROL;"
        Else
            str1 = "SELECT " & strSchema & ".VERSIONCONTROL.* "
            str2 = "FROM " & strSchema & ".VERSIONCONTROL;"
        End If

        strSQL = str1 & str2
        '''Console.WriteLine("tblAAUnkRunID: " & strSQL)

        Dim rsWVC As New ADODB.Recordset
        rsWVC.CursorLocation = CursorLocationEnum.adUseClient
        rsWVC.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        rsWVC.ActiveConnection = Nothing

        tblWatsonDBVersion.Clear()
        tblWatsonDBVersion.AcceptChanges()
        tblWatsonDBVersion.BeginLoadData()
        daDoPr.Fill(tblWatsonDBVersion, rsWVC)
        tblWatsonDBVersion.EndLoadData()
        rsWVC.Close()
        rsWVC = Nothing

        'Public strWatsonDBVersion As String
        'Public intWatsonDBVersion As Int16

        'record Watson database version
        Call SetvWatsonDB()


        '*****

        'but first traditional arrAnalytes and tblAnalyteHome must be prepared

        '*****begin arrAnalytes

        Try
            'ANARUNRAWANALYTEPEAK   OLD
            'ANARUNRAWANALYTEPEAK_INJECT   NEW
            'Dim strAnaRunPeak As String 'this is global "ANARUNRAWANALYTEPEAK"
            If boolAccess Then
                Try
                    If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rs.Close()
                    End If
                    rs.Open("SELECT ANARUNRAWANALYTEPEAK.STUDYID FROM ANARUNRAWANALYTEPEAK WHERE ANARUNRAWANALYTEPEAK.STUDYID < 1", cn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
                    rs.Close()
                    strAnaRunPeak = "ANARUNRAWANALYTEPEAK"
                Catch ex As Exception
                    strAnaRunPeak = "ANARUNRAWANALYTEPEAK_INJECT"
                    strF = ex.Message
                    strF = strF
                End Try
            Else
                Try
                    If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rs.Close()
                    End If
                    rs.Open("SELECT " & strSchema & ".ANARUNRAWANALYTEPEAK.STUDYID FROM " & strSchema & ".ANARUNRAWANALYTEPEAK WHERE " & strSchema & ".ANARUNRAWANALYTEPEAK.STUDYID < 1", cn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
                    rs.Close()
                    strAnaRunPeak = "ANARUNRAWANALYTEPEAK"
                Catch ex As Exception
                    strAnaRunPeak = "ANARUNRAWANALYTEPEAK_INJECT"
                    strF = ex.Message
                    strF = strF
                End Try
            End If

            str1 = "Retrieving active analyte info..."

            ''!!!!!New sql that pulls only analytes used in study!!
            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, GLOBALANALYTES.ANALYTEDESCRIPTION, GLOBALANALYTES.PROJECTID, ASSAY.MASTERASSAYID "
                ' NDL 7-Jan-2015: Consider adding specific range and matrix differentiators (see below)
                ' str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, GLOBALANALYTES.ANALYTEDESCRIPTION, GLOBALANALYTES.PROJECTID, ASSAY.MASTERASSAYID, ASSAY.SAMPLETYPEKEY, ASSAYANALYTES.VEC, ASSAYANALYTES.NM
                str2 = "FROM ANARUNANALYTERESULTS INNER JOIN (ASSAY INNER JOIN ((ASSAYANALYTES INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID) INNER JOIN STUDY ON ASSAYANALYTES.STUDYID = STUDY.STUDYID) ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) ON ANARUNANALYTERESULTS.RUNID = ASSAY.RUNID "
                str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & wStudyID & ") And ((GLOBALANALYTES.ACTIVE) = -1) AND ((ASSAY.RUNID) > 0)) "
                str4 = "ORDER BY GLOBALANALYTES.ANALYTEDESCRIPTION;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".GLOBALANALYTES.PROJECTID, " & strSchema & ".ASSAY.MASTERASSAYID "
                str2 = "FROM " & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (" & strSchema & ".ASSAY INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN " & strSchema & ".GLOBALANALYTES ON " & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID) INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".STUDY.STUDYID) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON (" & strSchema & ".ANARUNANALYTERESULTS.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ASSAY.RUNID) "
                str3 = "WHERE(((" & strSchema & ".ASSAYANALYTES.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".GLOBALANALYTES.ACTIVE) = -1) AND ((" & strSchema & ".ASSAY.RUNID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION;"
            End If

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("rsAnalytes: " & strSQL)
            ' Debug.WriteLine("Analytes: " & strSQL)
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient

            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing
            int1 = rs.RecordCount 'debug

            'clear contents of arranalytes
            ReDim arrAnalytes(intUBAA, 500)

            Count1 = 0
            Count2 = 0
            Try
                Do Until rs.EOF  'For each Analyte Record
                    '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                    '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                    '10=UseIntStd, 11=IntStd, 12=MasterAssayID,13=IsCoadministeredCmpd,14=Original Analyte Description
                    var1 = rs.Fields("AnalyteDescription").Value
                    var2 = rs.Fields("MASTERASSAYID").Value
                    var3 = rs.Fields("AnalyteIndex").Value
                    strF = "MASTERASSAYID = " & var2 & " AND ANALYTEINDEX = " & var3
                    'rsC.Filter = ""
                    'rsC.Filter = strF

                    'NO! For Method Validation, there may be compounds included that have not been regressed, but need to have other stuff
                    'If rsC.RecordCount = 0 Then

                    Count1 = Count1 + 1
                    'determine if var1 already exists

                    'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                    '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                    '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET

                    arrAnalytes(1, Count1) = Replace(var1, ChrW(12288), " ", 1, -1, CompareMethod.Text) 'rs.Fields("AnalyteDescription").Value
                    arrAnalytes(2, Count1) = rs.Fields("AnalyteID").Value 'GlobalAnalyteID=AnalyteID
                    arrAnalytes(3, Count1) = rs.Fields("AnalyteIndex").Value
                    arrAnalytes(8, Count1) = "No"
                    arrAnalytes(9, Count1) = "No"
                    arrAnalytes(10, Count1) = "Yes"
                    arrAnalytes(11, Count1) = ""
                    arrAnalytes(12, Count1) = rs.Fields("MASTERASSAYID").Value
                    arrAnalytes(13, Count1) = False
                    var4 = rs.Fields("AnalyteDescription").Value
                    arrAnalytes(14, Count1) = rs.Fields("AnalyteDescription").Value 'original AnalyteDescription

                    arrAnalytes(17, Count1) = Count1

                    rs.MoveNext()
                Loop
            Catch ex As Exception
                var1 = ex.Message
            End Try

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            If rsC.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsC.Close()
            End If
            rsC = Nothing

            'Rename Analytes to add Duplicates
            Dim intRepeatCtr As Int16

            For Count2 = 1 To (Count1 - 1)  'Go through analytes
                intRepeatCtr = 1
                var3 = arrAnalytes(1, Count2)
                For Count3 = (Count2 + 1) To Count1 'And compare to all other analytes
                    If (StrComp(var3, arrAnalytes(1, Count3)) = 0) Then
                        If intRepeatCtr = 1 Then
                            '  arrAnalytes(1, Count2) = arrAnalytes(1, Count2) & "_C" & 1  'Name first analyte _C1 if there's a repeat
                        End If
                        intRepeatCtr = intRepeatCtr + 1
                        arrAnalytes(1, Count3) = arrAnalytes(1, Count3) & "_C" & intRepeatCtr  'Name the other repeats _C2, _C3, etc.
                    End If
                Next
            Next

            ctAnalytes = Count1 'record number of analytes

            ReDim arrAnalytesCB(ctAnalytes - 1)
            For Count1 = 0 To ctAnalytes - 1
                arrAnalytesCB(Count1) = arrAnalytes(1, Count1 + 1)
            Next


            str1 = "Retrieving Watson Data...5 " & ctPB
            str1 = str1 & ChrW(10) & "...If the study is large, this step may take a few moments..."
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()

            Cursor.Current = Cursors.WaitCursor

            System.Windows.Forms.Application.DoEvents()

            'retrieve active internalstandards in table ASSAYANALYTES using studyid
            Count2 = ctAnalytes 'need this to continue arrAnalytes counter
            str1 = "Retrieving active internal standard info..."

            '*****
            ctAnalytes_IS = 0
            int1 = 0

            For Count1 = 1 To ctAnalytes

                If boolAccess Then
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.STUDYID, ASSAY.MASTERASSAYID, " & strAnaRunPeak & ".ANALYTEINDEX, " & strAnaRunPeak & ".INTERNALSTDNAME "
                    str2 = "FROM ASSAYANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) "
                    str3 = "WHERE (((ASSAYANALYTES.ANALYTEID)=" & arrAnalytes(2, Count1) & ") AND ((ASSAYANALYTES.STUDYID)=" & wStudyID & ") AND ((ASSAY.MASTERASSAYID)=" & arrAnalytes(12, Count1) & ") AND ((" & strAnaRunPeak & ".ANALYTEINDEX)=" & arrAnalytes(3, Count1) & "));"
                    str4 = ""
                Else
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME "
                    str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) "
                    str3 = "WHERE (((" & strSchema & ".ASSAYANALYTES.ANALYTEID)=" & arrAnalytes(2, Count1) & ") AND ((" & strSchema & ".ASSAYANALYTES.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAY.MASTERASSAYID)=" & arrAnalytes(12, Count1) & ") AND ((" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX)=" & arrAnalytes(3, Count1) & "));"
                    str4 = ""
                End If

                strSQL = str1 & str2 & str3 & str4
                'Console.WriteLine("IS: " & strSQL)
                If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rs.Close()
                End If
                rs.CursorLocation = CursorLocationEnum.adUseClient

                Try
                    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1

                End Try
                rs.ActiveConnection = Nothing

                int2 = rs.RecordCount


                If int2 = 0 Then
                Else

                    'var1 = arrAnalytes(3, Count1)

                    var2 = rs.Fields("INTERNALSTDNAME").Value
                    var1 = Replace(NZ(var2, ""), ChrW(12288), " ", 1, -1, CompareMethod.Text)
                    'ensure IS name doesn't exist in array
                    boolI = True
                    Try
                        For Count3 = 1 To ctAnalytes * 2
                            If Count3 > UBound(arrAnalytes, 2) Then
                                Exit For
                            End If
                            var2 = NZ(arrAnalytes(1, Count3), "")
                            If StrComp(NZ(var1, ""), var2, CompareMethod.Text) = 0 Then 'IS already exists
                                boolI = False
                            End If
                        Next
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try

                    'ctAnalytes_IS = 0
                    If boolI Then
                        Try
                            'keep using Count1 from ctAnalytes counter
                            '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                            '4=BLQ, 5=AQL, 6=ConcUnits, 7=AcceptedRuns
                            Count2 = Count2 + 1
                            int1 = int1 + 1

                            ctAnalytes_IS = int1 'Count1 'record number of Int Stds

                            var1 = rs.Fields("INTERNALSTDNAME").Value 'debug
                            arrAnalytes(1, Count2) = Replace(var1, ChrW(12288), " ", 1, -1, CompareMethod.Text) ' rs.Fields("INTERNALSTDNAME").Value
                            'arrAnalytes(2, Count1) = rs.Fields("GlobalAnalyteID").Value 'GlobalAnalyteID=AnalyteID

                            'var1 = rs.Fields("ANALYTEINDEX") 'debug
                            'arrAnalytes(3, Count2) = rs.Fields("ANALYTEINDEX").Value 'NO!! IntStd has no analyteindex!!!


                            'Sheets("AnalRefTables").Range("AnalyteName").Offset(0, Count2).Value = arrAnalytes(1, Count2)
                            'Sheets("AnalRefTables").Range("IsReplicate").Offset(0, Count2).Value = "No"
                            arrAnalytes(8, Count2) = "No"
                            'Sheets("AnalRefTables").Range("IsInternalStandard").Offset(0, Count2).Value = "Yes"
                            arrAnalytes(9, Count2) = "Yes"
                            'Sheets("AnalRefTables").Range("UseInternalStandard").Offset(0, Count1).Value = "Yes"
                            arrAnalytes(10, Count2) = "NA"
                            'Sheets("AnalRefTables").Range("InternalStandard").Offset(0, Count1).Value = arrAnalytes(1, Count2)
                            arrAnalytes(11, Count2) = "NA"

                            'var1 = rs.Fields("MASTERASSAYID").Value 'debug
                            'arrAnalytes(12, Count2) = rs.Fields("MASTERASSAYID").Value 'NO!! IntStd has no masterassayid!!!

                            arrAnalytes(13, Count2) = False ' is duplicate  :is coadministeredcmpd?

                            arrAnalytes(17, Count2) = Count2
                        Catch ex As Exception
                            var1 = ex.Message
                        End Try

                    Else

                    End If
                    'If boolI Then
                    '    arrAnalytes(11, Count1) = arrAnalytes(1, Count2) 'record IS name
                    'End If
                    Try
                        arrAnalytes(11, Count1) = NZ(var1, "") ' arrAnalytes(1, Count2) 'record IS name
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try

                    'If ctAnalytes_IS = 0 Then
                    'Else
                    '    arrAnalytes(11, Count1) = arrAnalytes(1, Count2) 'record IS name
                    'End If
                End If

                If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rs.Close()
                End If
            Next
            ''*****

            ReDim Preserve arrAnalytes(intUBAA, ctAnalytes + ctAnalytes_IS)

            Dim drow8 As DataRow
            'save this array in a datatable
            'add columns to table tblanalyteshome
            'Note: Do not change the order of column addition
            'Custom reports are dependent on this column order
            tblAnalytesHome.Clear()
            tblAnalytesHome.AcceptChanges()
            If tblAnalytesHome.Columns.Count > 0 Then
            Else
                For Count1 = 1 To intUBAA
                    Select Case Count1
                        Case 1
                            str1 = "AnalyteDescription"
                            var1 = System.Type.GetType("System.String")
                        Case 2
                            str1 = "AnalyteID"
                            var1 = System.Type.GetType("System.Int64")
                        Case 3
                            str1 = "AnalyteIndex"
                            var1 = System.Type.GetType("System.Int64")
                        Case 4
                            str1 = "BQL"
                            var1 = System.Type.GetType("System.Single")
                        Case 5
                            str1 = "AQL"
                            var1 = System.Type.GetType("System.Single")
                        Case 6
                            str1 = "ConcUnits"
                            var1 = System.Type.GetType("System.String")
                        Case 7
                            str1 = "AcceptedRuns"
                            var1 = System.Type.GetType("System.Int64")
                        Case 8
                            str1 = "IsReplicate"
                            var1 = System.Type.GetType("System.String")
                        Case 9
                            str1 = "IsIntStd"
                            var1 = System.Type.GetType("System.String")
                        Case 10
                            str1 = "UseIntStd"
                            var1 = System.Type.GetType("System.String")
                        Case 11
                            str1 = "IntStd"
                            var1 = System.Type.GetType("System.String")
                        Case 12
                            str1 = "MasterAssayID"
                            var1 = System.Type.GetType("System.Int64")
                        Case 13
                            str1 = "IsCoadminCmpd"
                            var1 = System.Type.GetType("System.String")
                        Case 14
                            str1 = "ORIGINALANALYTEDESCRIPTION"
                            var1 = System.Type.GetType("System.String")
                        Case 15
                            str1 = "INTGROUP"
                            var1 = System.Type.GetType("System.Int16")
                        Case 16
                            str1 = "MATRIX"
                            var1 = System.Type.GetType("System.String")
                        Case 17
                            str1 = "INTORDER"
                            var1 = System.Type.GetType("System.Int16")
                        Case 18
                            str1 = "CALIBRSET"
                            var1 = System.Type.GetType("System.String")

                        Case 19
                            str1 = "CHARUSERANALYTE"
                            var1 = System.Type.GetType("System.String")
                        Case 20
                            str1 = "CHARUSERIS"
                            var1 = System.Type.GetType("System.String")

                    End Select

                    '20190206 LEE:
                    Try
                        Dim col As New DataColumn
                        col.ColumnName = str1
                        col.DataType = var1
                        col.AllowDBNull = True
                        tblAnalytesHome.Columns.Add(col)
                    Catch ex As Exception
                        var1 = var1
                    End Try
             
                Next
            End If

            For Count1 = 1 To ctAnalytes + ctAnalytes_IS
                drow8 = tblAnalytesHome.NewRow
                'intUBAA
                'For Count2 = 1 To 18 'Note: don't add groups here:15
                For Count2 = 1 To intUBAA
                    str1 = ""
                    Select Case Count2
                        Case 1
                            str1 = "AnalyteDescription"
                        Case 2
                            str1 = "AnalyteID"
                        Case 3
                            str1 = "AnalyteIndex"
                        Case 4
                            str1 = "BQL"
                        Case 5
                            str1 = "AQL"
                        Case 6
                            str1 = "ConcUnits"
                        Case 7
                            str1 = "AcceptedRuns"
                        Case 8
                            str1 = "IsReplicate"
                        Case 9
                            str1 = "IsIntStd"
                        Case 10
                            str1 = "UseIntStd"
                        Case 11
                            str1 = "IntStd"
                        Case 12
                            str1 = "MasterAssayID"
                        Case 13
                            str1 = "IsCoadminCmpd"
                        Case 14
                            str1 = "ORIGINALANALYTEDESCRIPTION"
                        Case 16
                            str1 = "MATRIX"
                        Case 17
                            str1 = "INTORDER"
                        Case 18
                            str1 = "CALIBRSET"

                        Case 19
                            str1 = "CHARUSERANALYTE"
                        Case 20
                            str1 = "CHARUSERIS"

                    End Select
                    If Len(str1) = 0 Then
                    Else
                        '20190206 LEE:
                        Try
                            var2 = arrAnalytes(Count2, Count1) 'debug
                            var1 = NZ(arrAnalytes(Count2, Count1), System.DBNull.Value)
                            '20190205 LEE:
                            'Hmmm. This may throw an error
                            Try
                                drow8(str1) = var1
                            Catch ex As Exception
                                var1 = var1
                            End Try
                        Catch ex As Exception
                            var1 = var1
                        End Try

                    End If
                Next
                Try
                    tblAnalytesHome.Rows.Add(drow8)
                Catch ex As Exception
                    var1 = var1
                End Try

            Next

            'Note: do note redim arranalytes
            'Later code will evaluate > ctAnalytes+ctAnalytes_IS
            'ReDim Preserve arrAnalytes(intUBAA, ctAnalytes + ctAnalytes_IS)

            '*****end arrAnalytes

            'tblAllStdsAssay (All Standards and QC's for all runs in One table with Nominal Concentrations, etc)
            'NDL 4-Feb-2016 Note that this is a *temporary* table, as it includes non-Accepted runs (including rejected runs)
            'We filter out the rejected runs in the specific tblBCStdsAssayID and tblBCQCStdsAssayID queries


            ''****

            If boolAccess Then

                str1 = "SELECT ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, GLOBALANALYTES.ANALYTEDESCRIPTION, ASSAYANALYTES.ASSAYID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.RUNSTARTDATE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAYANALYTES.INTERNALSTANDARD, ASSAYANALYTES.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAY.MASTERASSAYID "
                str2 = "FROM (GLOBALANALYTES INNER JOIN ((ASSAYANALYTES INNER JOIN ((ANALYTICALRUN INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, ASSAYANALYTEKNOWN.LEVELNUMBER"

                '20160310 LEE:
                'This query needs to have ASSAYREPS.ID for EvalOutliers
                str1 = "SELECT ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, GLOBALANALYTES.ANALYTEDESCRIPTION, ASSAYANALYTES.ASSAYID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.RUNSTARTDATE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAYANALYTES.INTERNALSTANDARD, ASSAYANALYTES.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAY.MASTERASSAYID, ASSAYREPS.ID "
                str2 = "FROM ((GLOBALANALYTES INNER JOIN ((ASSAYANALYTES INNER JOIN ((ANALYTICALRUN INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ASSAYREPS ON (ASSAYANALYTEKNOWN.STUDYID = ASSAYREPS.STUDYID) AND (ASSAYANALYTEKNOWN.LEVELNUMBER = ASSAYREPS.LEVELNUMBER) AND (ASSAYANALYTEKNOWN.KNOWNTYPE = ASSAYREPS.KNOWNTYPE) AND (ASSAYANALYTEKNOWN.ASSAYID = ASSAYREPS.ASSAYID) "
                str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160319 LEE: added , CONFIGRUNTYPES.RUNTYPEDESCRIPTION
                str1 = "SELECT ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, GLOBALANALYTES.ANALYTEDESCRIPTION, ASSAYANALYTES.ASSAYID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.RUNSTARTDATE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAYANALYTES.INTERNALSTANDARD, ASSAYANALYTES.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAY.MASTERASSAYID, ASSAYREPS.ID, CONFIGRUNTYPES.RUNTYPEDESCRIPTION "
                str2 = "FROM CONFIGRUNTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN ((ASSAYANALYTES INNER JOIN ((ANALYTICALRUN INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ASSAYREPS ON (ASSAYANALYTEKNOWN.ASSAYID = ASSAYREPS.ASSAYID) AND (ASSAYANALYTEKNOWN.KNOWNTYPE = ASSAYREPS.KNOWNTYPE) AND (ASSAYANALYTEKNOWN.LEVELNUMBER = ASSAYREPS.LEVELNUMBER) AND (ASSAYANALYTEKNOWN.STUDYID = ASSAYREPS.STUDYID)) ON CONFIGRUNTYPES.RUNTYPEID = ANALYTICALRUN.RUNTYPEID "
                str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160911 LEE: added , ASSAYREPS.FLAGPERCENT because sometimes ANALYTEFLAGPERCENT is blank
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, GLOBALANALYTES.ANALYTEDESCRIPTION, ASSAYANALYTES.ASSAYID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.RUNSTARTDATE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAYANALYTES.INTERNALSTANDARD, ASSAYANALYTES.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAY.MASTERASSAYID, ASSAYREPS.ID, CONFIGRUNTYPES.RUNTYPEDESCRIPTION, ASSAYREPS.FLAGPERCENT "
                str2 = "FROM CONFIGRUNTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN ((ASSAYANALYTES INNER JOIN ((ANALYTICALRUN INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ASSAYREPS ON (ASSAYANALYTEKNOWN.ASSAYID = ASSAYREPS.ASSAYID) AND (ASSAYANALYTEKNOWN.KNOWNTYPE = ASSAYREPS.KNOWNTYPE) AND (ASSAYANALYTEKNOWN.LEVELNUMBER = ASSAYREPS.LEVELNUMBER) AND (ASSAYANALYTEKNOWN.STUDYID = ASSAYREPS.STUDYID)) ON CONFIGRUNTYPES.RUNTYPEID = ANALYTICALRUN.RUNTYPEID "
                str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.RUNID, ASSAYANALYTEKNOWN.LEVELNUMBER;"

            Else

                str1 = "SELECT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " _
                    & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " _
                    & strSchema & ".ASSAYANALYTES.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " _
                    & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " _
                    & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " _
                    & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAYANALYTES.INTERNALSTANDARD, " _
                    & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " _
                    & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX "
                str2 = "FROM (" & strSchema & ".GLOBALANALYTES INNER JOIN ((" _
                    & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUN INNER JOIN " _
                    & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " _
                    & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " _
                    & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAY ON (" _
                    & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" _
                    & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" _
                    & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" _
                    & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" _
                    & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & "..ANALYTEINDEX)) INNER JOIN " _
                    & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " _
                    & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " _
                    & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " _
                    & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " _
                    & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " _
                    & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE(((" & strSchema & ".ASSAYANALYTES.studyid) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " _
                    & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER"

                '20160310 LEE:
                'This query needs to have ASSAYREPS.ID for EvalOutliers
                str1 = "SELECT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ASSAYANALYTES.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAYANALYTES.INTERNALSTANDARD, " & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYREPS.ID "
                str2 = "FROM ((" & strSchema & ".GLOBALANALYTES INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ASSAYREPS ON (" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID = " & strSchema & ".ASSAYREPS.STUDYID) AND (" & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER = " & strSchema & ".ASSAYREPS.LEVELNUMBER) AND (" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE = " & strSchema & ".ASSAYREPS.KNOWNTYPE) AND (" & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAYREPS.ASSAYID) "
                str3 = "WHERE(((" & strSchema & ".ASSAYANALYTES.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160319 LEE: added , CONFIGRUNTYPES.RUNTYPEDESCRIPTION
                str1 = "SELECT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ASSAYANALYTES.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAYANALYTES.INTERNALSTANDARD, " & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".CONFIGRUNTYPES.RUNTYPEDESCRIPTION "
                str2 = "FROM " & strSchema & ".CONFIGRUNTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID =" & strSchema & ". ASSAY.RUNID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ASSAYREPS ON (" & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAYREPS.ASSAYID) AND (" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE = " & strSchema & ".ASSAYREPS.KNOWNTYPE) AND (" & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER = " & strSchema & ".ASSAYREPS.LEVELNUMBER) AND (" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID = " & strSchema & ".ASSAYREPS.STUDYID)) ON " & strSchema & ".CONFIGRUNTYPES.RUNTYPEID = " & strSchema & ".ANALYTICALRUN.RUNTYPEID "
                str3 = "WHERE(((" & strSchema & ".ASSAYANALYTES.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160911 LEE: added , ASSAYREPS.FLAGPERCENT because sometimes ANALYTEFLAGPERCENT is blank
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ASSAYANALYTES.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAYANALYTES.INTERNALSTANDARD, " & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".CONFIGRUNTYPES.RUNTYPEDESCRIPTION, " & strSchema & ".ASSAYREPS.FLAGPERCENT "
                str2 = "FROM " & strSchema & ".CONFIGRUNTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID =" & strSchema & ". ASSAY.RUNID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ASSAYREPS ON (" & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAYREPS.ASSAYID) AND (" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE = " & strSchema & ".ASSAYREPS.KNOWNTYPE) AND (" & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER = " & strSchema & ".ASSAYREPS.LEVELNUMBER) AND (" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID = " & strSchema & ".ASSAYREPS.STUDYID)) ON " & strSchema & ".CONFIGRUNTYPES.RUNTYPEID = " & strSchema & ".ANALYTICALRUN.RUNTYPEID "
                str3 = "WHERE(((" & strSchema & ".ASSAYANALYTES.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

            End If

            ', ASSAYREPS.FLAGPERCENT

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblAllStdsAssayID: " & strSQL)

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing
            If rs.EOF And rs.BOF Then
            Else
                rs.MoveFirst()
            End If


            'save this recordset in a datatable
            tblAllStdsAssay.Clear()
            tblAllStdsAssay.AcceptChanges()
            tblAllStdsAssay.BeginLoadData()
            daDoPr.Fill(tblAllStdsAssay, rs)
            tblAllStdsAssay.EndLoadData()
            rs.Close()



            'debug
            'Dim intAAAA As Int64
            'intAAAA = tblAllStdsAssay.Rows.Count
            'MsgBox(intAAAA)


            Call FixTableConcentrations(tblAllStdsAssay)

            'this routine requires the following tables to be populated beforehand
            'tblBCStdsAssayID

            Dim dv5 As New DataView(tblAllStdsAssay)
            strF = "(KNOWNTYPE='STANDARD') AND (RUNANALYTEREGRESSIONSTATUS=3) AND (RUNTYPEID<>3)"
            dv5.RowFilter = strF
            'tblBCStdsAssayID = dv5.ToTable("tblBCStdsAssayID", True, "MASTERASSAYID", "ANALYTEINDEX", _
            '                               "ANALYTEID", "LEVELNUMBER", "CONCENTRATION", "STUDYID", "ASSAYID", _
            '                               "ANALYTEFLAGPERCENT", "RUNID", "RUNTYPEID", "INTERNALSTANDARD", "SAMPLETYPEID")
            tblBCStdsAssayID = dv5.ToTable
            ''


            'tblBCQCStdsAssayID

            strF = "(KNOWNTYPE='QC') AND (RUNANALYTEREGRESSIONSTATUS=3) AND (RUNTYPEID<>3)"
            dv5.RowFilter = strF
            'tblBCQCStdsAssayID = dv5.ToTable("tblBCStdsAssayID", True, "MASTERASSAYID", "ANALYTEINDEX", _
            '                               "ANALYTEID", "LEVELNUMBER", "CONCENTRATION", "STUDYID", "ASSAYID", _
            '                               "ANALYTEFLAGPERCENT", "RUNID", "RUNTYPEID", "INTERNALSTANDARD", "SAMPLETYPEID")
            tblBCQCStdsAssayID = dv5.ToTable()



            '*****

            'tblAccAnalRuns

            'rsAAR retrieves accepted calibration std info filtered for accepted analyticatl runs
            If boolAccess Then
                'str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID "
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID, ASSAY.ASSAYID "
                str2 = "FROM ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) "
                str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) <> 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX;"
            Else
                If boolANSI Then
                    'str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID "
                    str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.STUDYID, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID "
                    str2 = "FROM " & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) "
                    str3 = "WHERE(((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) <> 3)) "
                    str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX;"
                Else
                    'str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID "
                    str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID, ASSAY.ASSAYID "
                    str2 = "FROM ASSAY, ANALYTICALRUN, ANALYTICALRUNANALYTES, ANARUNREGPARAMETERS "
                    str2 = str2 & "WHERE (((ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX)) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) "
                    str3 = "AND (((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) <> 3)) "
                    str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX;"
                End If
            End If

            ', ANALYTICALRUN.RUNTYPEID
            '20150826 Larry: added RUNTYPEID to differentiate between:
            '  CONFIGRUNTYPES:
            '   RUNTYPEDESCRIPTION	RUNTYPEID
            '   UNKNOWNS	        1
            '   VALIDATION	        2
            '   PSAE	            3
            '   MANDATORY REPEATS	4
            '   RECOVERY	        5

            If boolAccess Then
                ''str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID "
                'str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID "
                'str2 = "FROM ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) "
                'str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) <> 3)) "
                'str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX;"


                ''20160125 LEE: from Nick: joins to get AnalyteID
                'str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                'str2 = "FROM GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID "
                'str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<>3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3));"
                'str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                ''20160126 LEE:  Added Matrix (SAMPLETYPEID)
                'str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID "
                'str2 = "FROM (GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                'str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<>3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                'str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                '20160126 LEE:  Added NM and VEC
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.NM, ASSAYANALYTES.VEC "
                str2 = "FROM (GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<>3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                ', ASSAYANALYTES.NM, ASSAYANALYTES.VEC
            Else
                If boolANSI Then
                    ''str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID "
                    'str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.STUDYID, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID "
                    'str2 = "FROM " & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) "
                    'str3 = "WHERE(((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) <> 3)) "
                    'str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX;"

                    ''20160125 LEE: from Nick: joins to get AnalyteID
                    'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                    'str2 = "FROM " & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                    'str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3) AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3));"
                    'str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    ''20160126 LEE:  Added Matrix (SAMPLETYPEID)
                    'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                    'str2 = "FROM (" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    'str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3) AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                    'str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    '20160126 LEE:  Added NM and VEC
                    str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC "
                    str2 = "FROM (" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3) AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    ', ASSAYANALYTES.NM, ASSAYANALYTES.VEC

                Else

                End If
            End If

            'analyteid
            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblAccAnalRuns: " & strSQL)
            'rs1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'used to get min R2
            '20150911 Larry: rename to more descriptive name
            Dim rsAAR As New ADODB.Recordset
            If rsAAR.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsAAR.Close()
            End If
            rsAAR.CursorLocation = CursorLocationEnum.adUseClient
            rsAAR.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsAAR.ActiveConnection = Nothing

            tblAccAnalRuns.Clear()
            tblAccAnalRuns.AcceptChanges()
            tblAccAnalRuns.BeginLoadData()
            daDoPr.Fill(tblAccAnalRuns, rsAAR)
            tblAccAnalRuns.EndLoadData()


            'now do all analytical runs

            If boolAccess Then

                ''20160125 LEE: from Nick: joins to get AnalyteID
                'str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                'str2 = "FROM GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID "
                'str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                'str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                ''20160126 LEE:  Added Matrix (SAMPLETYPEID)
                'str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID "
                'str2 = "FROM (GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                'str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                'str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                '20160126 LEE:  Added NM and VEC
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.NM, ASSAYANALYTES.VEC "
                str2 = "FROM (GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                '20160205 LEE:  Added Conc Units
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.NM, ASSAYANALYTES.VEC, CONCENTRATIONUNITS.CONCENTRATIONUNITS "
                str2 = "FROM ((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONCENTRATIONUNITS ON ASSAYANALYTES.CONCUNITSID = CONCENTRATIONUNITS.CONCUNITSID "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                '20160319 LEE added , CONFIGRUNTYPES.RUNTYPEDESCRIPTION
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.NM, ASSAYANALYTES.VEC, CONCENTRATIONUNITS.CONCENTRATIONUNITS, CONFIGRUNTYPES.RUNTYPEDESCRIPTION "
                str2 = "FROM (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONCENTRATIONUNITS ON ASSAYANALYTES.CONCUNITSID = CONCENTRATIONUNITS.CONCUNITSID) INNER JOIN CONFIGRUNTYPES ON ANALYTICALRUN.RUNTYPEID = CONFIGRUNTYPES.RUNTYPEID "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                '20160715 LEE: added , ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.NM, ASSAYANALYTES.VEC, CONCENTRATIONUNITS.CONCENTRATIONUNITS, CONFIGRUNTYPES.RUNTYPEDESCRIPTION, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE "
                str2 = "FROM (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONCENTRATIONUNITS ON ASSAYANALYTES.CONCUNITSID = CONCENTRATIONUNITS.CONCUNITSID) INNER JOIN CONFIGRUNTYPES ON ANALYTICALRUN.RUNTYPEID = CONFIGRUNTYPES.RUNTYPEID "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                '20160729 LEE: added , ANALYTICALRUN.RUNDESCRIPTION
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.NM, ASSAYANALYTES.VEC, CONCENTRATIONUNITS.CONCENTRATIONUNITS, CONFIGRUNTYPES.RUNTYPEDESCRIPTION, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE, ANALYTICALRUN.RUNDESCRIPTION "
                str2 = "FROM (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONCENTRATIONUNITS ON ASSAYANALYTES.CONCUNITSID = CONCENTRATIONUNITS.CONCUNITSID) INNER JOIN CONFIGRUNTYPES ON ANALYTICALRUN.RUNTYPEID = CONFIGRUNTYPES.RUNTYPEID "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

            Else
                If boolANSI Then
                    ''20160125 LEE: from Nick: joins to get AnalyteID
                    'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                    'str2 = "FROM " & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                    'str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                    'str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    ''20160126 LEE:  Added Matrix (SAMPLETYPEID)
                    'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                    'str2 = "FROM (" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    'str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                    'str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    '20160126 LEE:  Added NM and VEC
                    str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID', " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC "
                    str2 = "FROM (" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    '20160205 LEE:  Added Conc Units
                    str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION," & strSchema & ". ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.NM, ASSAYANALYTES.VEC, " & strSchema & ".CONCENTRATIONUNITS.CONCENTRATIONUNITS "
                    str2 = "FROM ((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONCENTRATIONUNITS ON " & strSchema & ".ASSAYANALYTES.CONCUNITSID = " & strSchema & ".CONCENTRATIONUNITS.CONCUNITSID "
                    str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    '20160319 LEE added , CONFIGRUNTYPES.RUNTYPEDESCRIPTION
                    str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC, " & strSchema & ".CONCENTRATIONUNITS.CONCENTRATIONUNITS, " & strSchema & ".CONFIGRUNTYPES.RUNTYPEDESCRIPTION "
                    str2 = "FROM (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONCENTRATIONUNITS ON " & strSchema & ".ASSAYANALYTES.CONCUNITSID = " & strSchema & ".CONCENTRATIONUNITS.CONCUNITSID) INNER JOIN " & strSchema & ".CONFIGRUNTYPES ON " & strSchema & ".ANALYTICALRUN.RUNTYPEID = " & strSchema & ".CONFIGRUNTYPES.RUNTYPEID "
                    str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    '20160715 LEE: added , ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE
                    str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC, " & strSchema & ".CONCENTRATIONUNITS.CONCENTRATIONUNITS, " & strSchema & ".CONFIGRUNTYPES.RUNTYPEDESCRIPTION, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".ANALYTICALRUN.EXTRACTIONDATE "
                    str2 = "FROM (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONCENTRATIONUNITS ON " & strSchema & ".ASSAYANALYTES.CONCUNITSID = " & strSchema & ".CONCENTRATIONUNITS.CONCUNITSID) INNER JOIN " & strSchema & ".CONFIGRUNTYPES ON " & strSchema & ".ANALYTICALRUN.RUNTYPEID = " & strSchema & ".CONFIGRUNTYPES.RUNTYPEID "
                    str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                    '20160729 LEE: added , ANALYTICALRUN.RUNDESCRIPTION
                    str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC, " & strSchema & ".CONCENTRATIONUNITS.CONCENTRATIONUNITS, " & strSchema & ".CONFIGRUNTYPES.RUNTYPEDESCRIPTION, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".ANALYTICALRUN.EXTRACTIONDATE, " & strSchema & ".ANALYTICALRUN.RUNDESCRIPTION "
                    str2 = "FROM (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONCENTRATIONUNITS ON " & strSchema & ".ASSAYANALYTES.CONCUNITSID = " & strSchema & ".CONCENTRATIONUNITS.CONCUNITSID) INNER JOIN " & strSchema & ".CONFIGRUNTYPES ON " & strSchema & ".ANALYTICALRUN.RUNTYPEID = " & strSchema & ".CONFIGRUNTYPES.RUNTYPEID "
                    str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                Else
                End If
            End If

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblAllAnalRuns: " & strSQL)

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            tblAllAnalRuns.Clear()
            tblAllAnalRuns.AcceptChanges()
            tblAllAnalRuns.BeginLoadData()
            daDoPr.Fill(tblAllAnalRuns, rs)
            tblAllAnalRuns.EndLoadData()


            'tblAAUnkRunID

            'do tblAAUnkRunID
            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.ASSAYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTES.NM, ASSAYANALYTES.VEC, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNSTARTDATE "
                str2 = "FROM ANALYTICALRUN INNER JOIN (ASSAYANALYTES INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID)) ON (ANALYTICALRUN.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD' Or (ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTEKNOWN.ASSAYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE "
                str2 = "FROM " & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD' Or (" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"
            End If

            strSQL = str1 & str2 & str3 & str4
            ''Console.WriteLine("tblAAUnkRunID: " & strSQL)

            Dim rsAAUnkRunID As New ADODB.Recordset
            rsAAUnkRunID.CursorLocation = CursorLocationEnum.adUseClient
            rsAAUnkRunID.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsAAUnkRunID.ActiveConnection = Nothing

            tblAAUnkRunID.Clear()
            tblAAUnkRunID.AcceptChanges()
            tblAAUnkRunID.BeginLoadData()
            daDoPr.Fill(tblAAUnkRunID, rsAAUnkRunID)
            tblAAUnkRunID.EndLoadData()
            rsAAUnkRunID.Close()
            rsAAUnkRunID = Nothing


            '20160418 LEE:
            'added tblASSAYREPS for use in function FindLabelHelper1
            'FindLabelHelper1 only looks at QCs with sample type QC
            'sometimes users make their own sample types
            'all sample types are included in Watson table ASSAYANALYTEKNOWN

            'Note: This same SQL gets called in AssignedSamples.ChangeStudy and function AssignedSamples.cbxStudyValidating, so if changes are made here, remember to change change in AssignedSamples.ChangeStudy and function AssignedSamples.cbxStudyValidating

            If boolAccess Then
                str1 = "SELECT ASSAYREPS.* "
                str2 = "FROM ASSAYREPS "
                str3 = "WHERE (((ASSAYREPS.STUDYID)=" & wStudyID & "));"
            Else
                str1 = "SELECT " & strSchema & ".ASSAYREPS.* "
                str2 = "FROM " & strSchema & ".ASSAYREPS "
                str3 = "WHERE (((" & strSchema & ".ASSAYREPS.STUDYID)=" & wStudyID & "));"
            End If

            strSQL = str1 & str2 & str3
            'Console.WriteLine("tblASSAYREPS: " & strSQL)

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            Try
                rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                var1 = var1
            End Try

            rs.ActiveConnection = Nothing

            'rsspeciesmatrix
            tblASSAYREPS.Clear()
            tblASSAYREPS.AcceptChanges()
            tblASSAYREPS.BeginLoadData()
            daDoPr.Fill(tblASSAYREPS, rs)
            tblASSAYREPS.EndLoadData()


            '20160419 LEE:
            'pulled this table out of AssignSamples, doesn't need to get called so often
            'Note: This same SQL gets called in AssignedSamples.ChangeStudy and function AssignedSamples.cbxStudyValidating, so if changes are made here, remember to change change in AssignedSamples.ChangeStudy and function AssignedSamples.cbxStudyValidating
            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.STUDYID "
                str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER)) INNER JOIN GLOBALANALYTES ON (ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID)"
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
                str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) INNER JOIN " & strSchema & ".GLOBALANALYTES ON (" & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID)"
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

            End If

            'TBLSTUDIES
            strSQL = str1 & str2 & str3 & str4
            '''Console.WriteLine(strSQL)
            '''''''''''''''''''''''''''''''''''''''''Console.WriteLine(strSQL)
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            Try
                rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                var1 = var1
            End Try
            rs.ActiveConnection = Nothing

            'rsspeciesmatrix
            tblAnalyteConcLevelsForAssay.Clear()
            tblAnalyteConcLevelsForAssay.AcceptChanges()
            tblAnalyteConcLevelsForAssay.BeginLoadData()
            daDoPr.Fill(tblAnalyteConcLevelsForAssay, rs)
            tblAnalyteConcLevelsForAssay.EndLoadData()


            '20160124 LEE:  tblSpecies must be executed before boolUseGroups

            'record matrix and sample volume and species and peakheight/area
            'peak: 1=height, 0=area

            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAY.STUDYID, ASSAY.STANDARDVOLUME, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.ACTIVE, CONFIGSPECIES.SPECIES, ASSAY.HEIGHTORAREA "
                str2 = "FROM (ASSAY INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONFIGSPECIES ON ASSAY.SPECIESID = CONFIGSPECIES.SPECIESID "
                str3 = "WHERE (((ASSAY.STUDYID)=" & wStudyID & ") AND ((CONFIGSAMPLETYPES.ACTIVE)=-1));"

                '20171111 LEE: STANDARDVOLUME is screwing things up
                'STANDARDVOLUME can be different for an assay for the same matrix
                'In multi-matrix studies (Alturas - IT001-101), can get multiple matrices
                'create a different table for STANDARDVOLUME
                str1 = "SELECT DISTINCT ASSAY.STUDYID, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.ACTIVE, CONFIGSPECIES.SPECIES, ASSAY.HEIGHTORAREA "
                str2 = "FROM (ASSAY INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONFIGSPECIES ON ASSAY.SPECIESID = CONFIGSPECIES.SPECIESID "
                str3 = "WHERE (((ASSAY.STUDYID)=" & wStudyID & ") AND ((CONFIGSAMPLETYPES.ACTIVE)=-1));"


            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.STUDYID, " & strSchema & ".ASSAY.STANDARDVOLUME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID," & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.ACTIVE, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".ASSAY.HEIGHTORAREA "
                str2 = "FROM (" & strSchema & ".ASSAY INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".ASSAY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID "
                str3 = "WHERE (((" & strSchema & ".ASSAY.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".CONFIGSAMPLETYPES.ACTIVE)=-1));"

                '20171111 LEE: STANDARDVOLUME is screwing things up
                'STANDARDVOLUME can be different for an assay for the same matrix
                'In multi-matrix studies (Alturas - IT001-101), can get multiple matrices
                'create a different table for STANDARDVOLUME
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.STUDYID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID," & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.ACTIVE, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".ASSAY.HEIGHTORAREA "
                str2 = "FROM (" & strSchema & ".ASSAY INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".ASSAY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID "
                str3 = "WHERE (((" & strSchema & ".ASSAY.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".CONFIGSAMPLETYPES.ACTIVE)=-1));"

            End If

            strSQL = str1 & str2 & str3
            'Console.WriteLine("tblSpeciesMatrix: " & strSQL)
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            '20180522 LEE: LI00041 has provided studies that do not have an active matrix (CONFIGSAMPLETYPES.ACTIVE)=-1).
            'If tblSpeciesMatrix.recordcount = 0, StudyDoc cannot process the data
            'Put in a check here
            If rs.RecordCount = 0 Then

                If boolAccess Then
                    str1 = "SELECT DISTINCT ASSAY.STUDYID, ASSAY.STANDARDVOLUME, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.ACTIVE, CONFIGSPECIES.SPECIES, ASSAY.HEIGHTORAREA "
                    str2 = "FROM (ASSAY INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONFIGSPECIES ON ASSAY.SPECIESID = CONFIGSPECIES.SPECIESID "
                    str3 = "WHERE (((ASSAY.STUDYID)=" & wStudyID & ") AND ((CONFIGSAMPLETYPES.ACTIVE)=-1));"

                    '20171111 LEE: STANDARDVOLUME is screwing things up
                    'STANDARDVOLUME can be different for an assay for the same matrix
                    'In multi-matrix studies (Alturas - IT001-101), can get multiple matrices
                    'create a different table for STANDARDVOLUME
                    str1 = "SELECT DISTINCT ASSAY.STUDYID, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.ACTIVE, CONFIGSPECIES.SPECIES, ASSAY.HEIGHTORAREA "
                    str2 = "FROM (ASSAY INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONFIGSPECIES ON ASSAY.SPECIESID = CONFIGSPECIES.SPECIESID "
                    str3 = "WHERE (((ASSAY.STUDYID)=" & wStudyID & "));"


                Else
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.STUDYID, " & strSchema & ".ASSAY.STANDARDVOLUME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID," & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.ACTIVE, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".ASSAY.HEIGHTORAREA "
                    str2 = "FROM (" & strSchema & ".ASSAY INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".ASSAY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID "
                    str3 = "WHERE (((" & strSchema & ".ASSAY.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".CONFIGSAMPLETYPES.ACTIVE)=-1));"

                    '20171111 LEE: STANDARDVOLUME is screwing things up
                    'STANDARDVOLUME can be different for an assay for the same matrix
                    'In multi-matrix studies (Alturas - IT001-101), can get multiple matrices
                    'create a different table for STANDARDVOLUME
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.STUDYID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID," & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.ACTIVE, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".ASSAY.HEIGHTORAREA "
                    str2 = "FROM (" & strSchema & ".ASSAY INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".ASSAY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID "
                    str3 = "WHERE (((" & strSchema & ".ASSAY.STUDYID)=" & wStudyID & "));"

                End If

                strSQL = str1 & str2 & str3
                'Console.WriteLine("tblSpeciesMatrix: " & strSQL)
                If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rs.Close()
                End If
                rs.CursorLocation = CursorLocationEnum.adUseClient
                rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                rs.ActiveConnection = Nothing


            End If


            'rsspeciesmatrix
            tblSpeciesMatrix.Clear()
            tblSpeciesMatrix.AcceptChanges()
            tblSpeciesMatrix.BeginLoadData()
            daDoPr.Fill(tblSpeciesMatrix, rs)
            tblSpeciesMatrix.EndLoadData()

            'find number of species
            Dim dvSP As System.Data.DataView = New DataView(tblSpeciesMatrix)
            Dim tblSP As System.Data.DataTable = dvSP.ToTable("aSP", True, "SPECIES")
            intNumSpecies = tblSP.Rows.Count

            'find number of matrixes
            tblSP = dvSP.ToTable("aSM", True, "SAMPLETYPEID")
            intNumMatrix = tblSP.Rows.Count
            gNumMatrix = intNumMatrix


            '20171111 LEE:
            'Get STANDARDVOLUME

            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAY.STUDYID, ASSAY.STANDARDVOLUME, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.ACTIVE, CONFIGSPECIES.SPECIES, ASSAY.HEIGHTORAREA "
                str2 = "FROM (ASSAY INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN CONFIGSPECIES ON ASSAY.SPECIESID = CONFIGSPECIES.SPECIESID "
                str3 = "WHERE (((ASSAY.STUDYID)=" & wStudyID & ") AND ((CONFIGSAMPLETYPES.ACTIVE)=-1));"

            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.STUDYID, " & strSchema & ".ASSAY.STANDARDVOLUME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID," & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.ACTIVE, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".ASSAY.HEIGHTORAREA "
                str2 = "FROM (" & strSchema & ".ASSAY INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".ASSAY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID "
                str3 = "WHERE (((" & strSchema & ".ASSAY.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".CONFIGSAMPLETYPES.ACTIVE)=-1));"

            End If

            strSQL = str1 & str2 & str3
            'Console.WriteLine(strSQL)
            'Console.WriteLine("tblSpeciesMatrixSV: " & strSQL)
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            tblSpeciesMatrixSV.Clear()
            tblSpeciesMatrixSV.AcceptChanges()
            tblSpeciesMatrixSV.BeginLoadData()
            daDoPr.Fill(tblSpeciesMatrixSV, rs)
            tblSpeciesMatrixSV.EndLoadData()

            '*****

            '****begin preparing sample tables

            Dim rsDesign As New ADODB.Recordset
            'If rbSampleReport.Checked Then
            'retrieve sample design for appropriate analyte
            'If boolGender Then
            If boolANSI Then
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNTREATMENT.TREATMENTKEY = DESIGNSUBJECTTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) "
                str3 = "WHERE (((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
            Else
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM DESIGNSUBJECTTREATMENT, DESIGNTREATMENT, SAMPLERESULTS, DESIGNSUBJECT, DESIGNSAMPLE, DESIGNSUBJECTGROUP "
                str2 = str2 & "WHERE (((((DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) AND DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) AND (DESIGNTREATMENT.TREATMENTKEY = DESIGNSUBJECTTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) "
                str3 = "AND (((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
            End If

            str2 = "OLD FROM DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNTREATMENT.TREATMENTKEY = DESIGNSUBJECTTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) "
            str2 = "NEW FROM (DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "

            'add sampletypeid (matrix)
            If boolANSI Then
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str2 = "FROM (DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
            Else
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM DESIGNSUBJECTTREATMENT, DESIGNTREATMENT, SAMPLERESULTS, DESIGNSUBJECT, DESIGNSAMPLE, DESIGNSUBJECTGROUP, CONFIGSAMPLETYPES "
                str2 = str2 & "WHERE ((((((DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) AND DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) AND DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "AND (((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
            End If

            'add RUNSAMPLEORDERNUMBER
            str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER "
            str2 = "FROM ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
            str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
            str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

            'add SAMPLENAME
            str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME "
            str2 = "FROM " & strAnaRunPeak & " INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) "
            str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
            str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

            'add STUDYID
            str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID "
            str2 = "FROM " & strAnaRunPeak & " INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID)) ON (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
            str3 = "WHERE(((DESIGNSAMPLE.STUDYID) =  " & wStudyID & ")) "
            str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

            'added REPORTEDCONC
            If boolAccess Then
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC "
                str2 = "FROM " & strAnaRunPeak & " INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

                '20160220 LEE: Added WEEK, VISITTEXT, USERSAMPLEID
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID "
                str2 = "FROM " & strAnaRunPeak & " INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID)) ON (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

                '20160301 LEE: Added , ANALYTICALRUNSAMPLE.ASSAYDATETIME
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME "
                str2 = "FROM " & strAnaRunPeak & " INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID)) ON (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

                strSQL = str1 & str2 & str3 & str4
                '''Console.WriteLine("tblSampleDesign_Old1: " & strSQL)

                '20160303 LEE: Nick's query: Made inner join to return all sample results
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME "
                str2 = "FROM " & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID)) ON (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

                '20160313 LEE: Added CONCENTRATIONSTATUS for concentration null values
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS "
                str2 = "FROM " & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID)) ON (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

                '20160531 LEE: need to filter DesignSubjectID > 0
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS "
                str2 = "FROM " & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID)) ON (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE (((DESIGNSUBJECT.DESIGNSUBJECTID)>0) AND ((DESIGNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"


                ', DESIGNSUBJECTGROUP.SUBJECTGROUPID

                '20160718 LEE: , DESIGNSUBJECTGROUP.SUBJECTGROUPID: added to get appropriate sorting Sample Conc table
                ', SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE to aid in identifying BQL/AQL
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS, DESIGNSUBJECTGROUP.SUBJECTGROUPID, SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE "
                str2 = "FROM " & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID)) ON (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE (((DESIGNSUBJECT.DESIGNSUBJECTID)>0) AND ((DESIGNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"


                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS, DESIGNSUBJECTGROUP.SUBJECTGROUPID, SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE, DESIGNSUBJECTTREATMENT.DOSEAMOUNT, DOSEUNITS.DOSEUNITSDESCRIPTION "
                str2 = "FROM DOSEUNITS INNER JOIN (" & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) "
                str2 = str2 & "INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND "
                str2 = str2 & "(ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID)) ON DOSEUNITS.DOSEUNITSID = DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE (((DESIGNSUBJECT.DESIGNSUBJECTID)>0) AND ((DESIGNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"

                'must be right-join on DoseUnits
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS, DESIGNSUBJECTGROUP.SUBJECTGROUPID, SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE, DESIGNSUBJECTTREATMENT.DOSEAMOUNT, DOSEUNITS.DOSEUNITSDESCRIPTION "
                str2 = "FROM DOSEUNITS RIGHT JOIN (" & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID)) ON DOSEUNITS.DOSEUNITSID = DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"


                'ANARUNRAWANALYTEPEAK_INJECT

                'DESIGNSUBJECTTREATMENT.DOSEAMOUNT,, DOSEUNITS.DOSEUNITSDESCRIPTION

                ''20170708 LEE: Added , DESIGNSAMPLE.COMMENTMEMO
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS, DESIGNSUBJECTGROUP.SUBJECTGROUPID, SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE, DESIGNSUBJECTTREATMENT.DOSEAMOUNT, DOSEUNITS.DOSEUNITSDESCRIPTION, DESIGNSAMPLE.COMMENTMEMO "
                str2 = "FROM DOSEUNITS RIGHT JOIN (" & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID)) ON DOSEUNITS.DOSEUNITSID = DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"


                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID,  Round([SAMPLERESULTS].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS, DESIGNSUBJECTGROUP.SUBJECTGROUPID, SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE, DESIGNSUBJECTTREATMENT.DOSEAMOUNT, DOSEUNITS.DOSEUNITSDESCRIPTION, DESIGNSAMPLE.COMMENTMEMO "
                str2 = "FROM DOSEUNITS RIGHT JOIN (" & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID)) ON DOSEUNITS.DOSEUNITSID = DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
                'WEEK


                '20180625 LEE:
                'Needed to add Year, Month from DESIGNEVENTSAMPLETIME
                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID,  Round([SAMPLERESULTS].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS, DESIGNSUBJECTGROUP.SUBJECTGROUPID, SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE, DESIGNSUBJECTTREATMENT.DOSEAMOUNT, DOSEUNITS.DOSEUNITSDESCRIPTION, DESIGNSAMPLE.COMMENTMEMO, DESIGNSUBJECTTREATMENT.Year, DESIGNSUBJECTTREATMENT.Month "
                str2 = "FROM DOSEUNITS RIGHT JOIN (" & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID)) ON DOSEUNITS.DOSEUNITSID = DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
                'assaydatetime
                'YEAR

                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANARUNRAWANALYTEPEAK.SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.Week, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, DESIGNSUBJECTTREATMENT.YEAR, DESIGNSUBJECTTREATMENT.MONTH, DESIGNSAMPLE.COMMENTMEMO "
                str2 = "FROM ANARUNRAWANALYTEPEAK INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANARUNRAWANALYTEPEAK.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANARUNRAWANALYTEPEAK.RUNID = SAMPLERESULTS.RUNID) AND (ANARUNRAWANALYTEPEAK.STUDYID = SAMPLERESULTS.STUDYID) "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.YEAR, DESIGNSUBJECTTREATMENT.MONTH, DESIGNSUBJECTTREATMENT.Week, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND;"
                'STARTSECOND
                'COMMENTMEMO

                str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID,  Round([SAMPLERESULTS].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strAnaRunPeak & ".SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT, DESIGNSAMPLE.USERSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, SAMPLERESULTS.CONCENTRATIONSTATUS, DESIGNSUBJECTGROUP.SUBJECTGROUPID, SAMPLERESULTS.CONCENTRATIONSTATUS, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CALIBRATIONRANGE, DESIGNSUBJECTTREATMENT.DOSEAMOUNT, DOSEUNITS.DOSEUNITSDESCRIPTION, DESIGNSAMPLE.COMMENTMEMO, DESIGNSUBJECTTREATMENT.YEAR, DESIGNSUBJECTTREATMENT.MONTH "
                str2 = "FROM DOSEUNITS RIGHT JOIN (" & strAnaRunPeak & " RIGHT JOIN (ANALYTICALRUNSAMPLE RIGHT JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = SAMPLERESULTS.STUDYID)) ON DOSEUNITS.DOSEUNITSID = DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
                'flag
                'year


            Else

                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC "
                str2 = "FROM " & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                '20160220 LEE: Added WEEK, VISITTEXT, USERSAMPLEID
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID "
                str2 = "FROM " & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                '20160301 LEE: Added , ANALYTICALRUNSAMPLE.ASSAYDATETIME
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME "
                str2 = "FROM " & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                '20160303 LEE: Nick's query: Made inner join to return all sample results
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME "
                str2 = "FROM " & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                '20160313 LEE: Added CONCENTRATIONSTATUS for concentration null values
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS "
                str2 = "FROM " & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                '20160531 LEE: need to filter DesignSubjectID > 0
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS "
                str2 = "FROM " & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE (((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)>0) AND ((" & strSchema & ".DESIGNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                '20160718 LEE: , DESIGNSUBJECTGROUP.SUBJECTGROUPID: added to get appropriate sorting Sample Conc table
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGE "
                str2 = "FROM " & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) "
                str3 = "WHERE (((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)>0) AND ((" & strSchema & ".DESIGNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"


                '20170111 LEE: Added , DESIGNSUBJECTTREATMENT.DOSEAMOUNT, DOSEUNITS.DOSEUNITSDESCRIPTION
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGE, " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEAMOUNT, " & strSchema & ".DOSEUNITS.DOSEUNITSDESCRIPTION "
                str2 = "FROM " & strSchema & ".DOSEUNITS INNER JOIN (" & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID)) "
                str2 = str2 & "INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID =" & strSchema & ". SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND "
                str2 = str2 & "(" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON " & strSchema & ".DOSEUNITS.DOSEUNITSID = " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE (((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)>0) AND ((" & strSchema & ".DESIGNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                'must be right-join on DoseUnits
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGE, " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEAMOUNT, " & strSchema & ".DOSEUNITS.DOSEUNITSDESCRIPTION "
                str2 = "FROM " & strSchema & ".DOSEUNITS RIGHT JOIN (" & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND "
                str2 = str2 & "(" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON " & strSchema & ".DOSEUNITS.DOSEUNITSID = " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                'strAnaRunPeak
                'DESIGNSUBJECTTREATMENT.DOSEAMOUNT(), DOSEUNITS.DOSEUNITSDESCRIPTION
                'CALIBRATIONRANGE


                ' '20170708 LEE: Added , DESIGNSAMPLE.COMMENTMEMO
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGE, " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEAMOUNT, " & strSchema & ".DOSEUNITS.DOSEUNITSDESCRIPTION, DESIGNSAMPLE.COMMENTMEMO "
                str2 = "FROM " & strSchema & ".DOSEUNITS RIGHT JOIN (" & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND "
                str2 = str2 & "(" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON " & strSchema & ".DOSEUNITS.DOSEUNITSID = " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"


                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, ROUND(" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR," & intDFDec & ") AS ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGE, " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEAMOUNT, " & strSchema & ".DOSEUNITS.DOSEUNITSDESCRIPTION, " & strSchema & ".DESIGNSAMPLE.COMMENTMEMO "
                str2 = "FROM " & strSchema & ".DOSEUNITS RIGHT JOIN (" & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND "
                str2 = str2 & "(" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON " & strSchema & ".DOSEUNITS.DOSEUNITSID = " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                ''20180622 LEE:
                ''Needed to add Year, Month
                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, ROUND(" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR," & intDFDec & ") AS ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGE, " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEAMOUNT, " & strSchema & ".DOSEUNITS.DOSEUNITSDESCRIPTION, " & strSchema & ".DESIGNSAMPLE.COMMENTMEMO, " & strSchema & ".DESIGNSUBJECTTREATMENT.Year, " & strSchema & ".DESIGNSUBJECTTREATMENT.Month, " & strSchema & ".DESIGNSAMPLE.COMMENTMEMO "
                str2 = "FROM " & strSchema & ".DOSEUNITS RIGHT JOIN (" & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND "
                str2 = str2 & "(" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON " & strSchema & ".DOSEUNITS.DOSEUNITSID = " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"

                str1 = "SELECT DISTINCT " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".SAMPLERESULTS.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, ROUND(" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR," & intDFDec & ") AS ALIQUOTFACTOR, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".DESIGNSAMPLE.STUDYID, (" & strSchema & ".SAMPLERESULTS.CONCENTRATION/" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".SAMPLERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGE, " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEAMOUNT, " & strSchema & ".DOSEUNITS.DOSEUNITSDESCRIPTION, " & strSchema & ".DESIGNSAMPLE.COMMENTMEMO, " & strSchema & ".DESIGNSUBJECTTREATMENT.Year, " & strSchema & ".DESIGNSUBJECTTREATMENT.Month "
                str2 = "FROM " & strSchema & ".DOSEUNITS RIGHT JOIN (" & strSchema & "." & strAnaRunPeak & " RIGHT JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE RIGHT JOIN ((" & strSchema & ".DESIGNSUBJECTTREATMENT INNER JOIN (" & strSchema & ".DESIGNTREATMENT INNER JOIN (" & strSchema & ".SAMPLERESULTS INNER JOIN ((" & strSchema & ".DESIGNSUBJECT INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) ON " & strSchema & ".DESIGNTREATMENT.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND "
                str2 = str2 & "(" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".SAMPLERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) ON " & strSchema & ".DOSEUNITS.DOSEUNITSID = " & strSchema & ".DESIGNSUBJECTTREATMENT.DOSEUNITSID "
                str3 = "WHERE(((" & strSchema & ".DESIGNSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR;"



            End If

            ', , DESIGNSUBJECTGROUP.SUBJECTGROUPID
            'week

            'ANARUNRAWANALYTEPEAK ..

            strSQL = str1 & str2 & str3 & str4

            'Console.WriteLine("tblSampleDesign: " & strSQL)

            rsDesign.CursorLocation = CursorLocationEnum.adUseClient
            rsDesign.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'Try
            '    rsDesign.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'Catch ex As Exception
            '    str1 = "SELECT DISTINCT DESIGNSUBJECT.GENDERID, SAMPLERESULTS.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.RUNID, DESIGNSAMPLE.DESIGNSAMPLEID, SAMPLERESULTS.ALIQUOTFACTOR, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.TIMETEXT, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANARUNRAWANALYTEPEAK_INJECT.SAMPLENAME, DESIGNSAMPLE.STUDYID, (SAMPLERESULTS.CONCENTRATION/SAMPLERESULTS.ALIQUOTFACTOR) AS REPORTEDCONC "
            '    str2 = "FROM ANARUNRAWANALYTEPEAK_INJECT INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ((DESIGNSUBJECTTREATMENT INNER JOIN (DESIGNTREATMENT INNER JOIN (SAMPLERESULTS INNER JOIN ((DESIGNSUBJECT INNER JOIN DESIGNSAMPLE ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSAMPLE.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSAMPLE.STUDYID) AND (DESIGNSUBJECT.DESIGNSUBJECTID = DESIGNSAMPLE.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) ON DESIGNTREATMENT.STUDYID = SAMPLERESULTS.STUDYID) ON (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY = DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ANALYTICALRUNSAMPLE.STUDYID = SAMPLERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = SAMPLERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER = SAMPLERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNID = SAMPLERESULTS.RUNID) AND (ANARUNRAWANALYTEPEAK_INJECT.STUDYID = SAMPLERESULTS.STUDYID) "
            '    str3 = "WHERE(((DESIGNSAMPLE.STUDYID) = " & wStudyID & ")) "
            '    str4 = "ORDER BY DESIGNSUBJECT.GENDERID, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR;"
            '    strSQL = str1 & str2 & str3 & str4
            '    rsDesign.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'End Try

            tblSampleDesign.Clear()
            tblSampleDesign.AcceptChanges()
            tblSampleDesign.BeginLoadData()
            daDoPr.Fill(tblSampleDesign, rsDesign)
            tblSampleDesign.EndLoadData()

            If tblSampleDesign.Columns.Contains("AnalyteDescription") Then
            Else
                Dim col10 As New DataColumn
                col10.ColumnName = "AnalyteDescription"
                col10.DataType = System.Type.GetType("System.String")
                tblSampleDesign.Columns.Add(col10)
            End If

            If tblSampleDesign.Columns.Contains("SERIALENDTIME") Then
            Else
                Dim col11 As New DataColumn
                col11.ColumnName = "SERIALENDTIME"
                col11.DataType = System.Type.GetType("System.Int64")
                tblSampleDesign.Columns.Add(col11)
            End If

            If tblSampleDesign.Columns.Contains("SERIALSTARTTIME") Then
            Else
                Dim col12 As New DataColumn
                col12.ColumnName = "SERIALSTARTTIME"
                col12.DataType = System.Type.GetType("System.Int64")
                tblSampleDesign.Columns.Add(col12)
            End If

            Count2 = -1

            Do Until rsDesign.EOF
                'Dim drow As New DataRow
                Count2 = Count2 + 1
                'drow = tblSampleDesign.NewRow
                'drow("Analyte") = "Analyte"
                'tblSampleDesign.Rows.Item(Count1).BeginEdit()
                tblSampleDesign.Rows.Item(Count2).BeginEdit()
                'tblSampleDesign.Rows.Item(Count2).Item("Analyte") = "Analyte"
                tblSampleDesign.Rows.Item(Count2).Item("AnalyteDescription") = "Analyte"
                'For Each fld In rsDesign.Fields
                '    If StrComp(fld.Name, "Analyte", CompareMethod.Text) = 0 Then
                '        drow(fld.Name) = "Analyte"
                '    Else
                '        drow(fld.Name) = fld.Value
                '    End If
                'Next
                'tblSampleDesign.Rows.Add(drow)

                'now calculate SERIALENDTIME
                vS = NZ(rsDesign.Fields("ENDSECOND").Value, 0)
                vM = NZ(rsDesign.Fields("ENDMINUTE").Value, 0)
                vH = NZ(rsDesign.Fields("ENDHOUR").Value, 0)
                vD = NZ(rsDesign.Fields("ENDDAY").Value, 0)

                'convert to seconds
                vS = vS * 1
                vM = vM * 60
                vH = vH * 60 * 60
                vD = vD * 60 * 60 * 24
                sDate = vS + vM + vH + vD

                tblSampleDesign.Rows(Count2).Item("SERIALENDTIME") = sDate

                'now calculate SERIALSTARTTIME
                vS = NZ(rsDesign.Fields("STARTSECOND").Value, 0)
                vM = NZ(rsDesign.Fields("STARTMINUTE").Value, 0)
                vH = NZ(rsDesign.Fields("STARTHOUR").Value, 0)
                vD = NZ(rsDesign.Fields("STARTDAY").Value, 0)

                'convert to seconds
                vS = vS * 1
                vM = vM * 60
                vH = vH * 60 * 60
                vD = vD * 60 * 60 * 24
                sDate = vS + vM + vH + vD

                tblSampleDesign.Rows(Count2).Item("SERIALSTARTTIME") = sDate
                'tblSampleDesign.Rows.Item(Count1).EndEdit()
                tblSampleDesign.Rows.Item(Count2).EndEdit()

                rsDesign.MoveNext()

            Loop

            rsDesign.Close()
            rsDesign = Nothing

            '*** End tblSampleDesign

            Try
                If boolUseGroups Then

                    'Note that tblAnalyteGroups gets created in EstablishCalStdGroups(true)

                    Try
                        str1 = "CreatetblCalStdGroups"
                        Call CreatetblCalStdGroups()
                        str1 = "CreatetblCalStdGroupAssayIDs"
                        Call CreatetblCalStdGroupAssayIDs()

                    Catch ex As Exception
                        var1 = ex.Message
                    End Try

                    'Note that tblAnalyteGroups gets created in EstablishCalStdGroups

                    Try
                        Call EstablishCalStdGroups(True, cn) 'cal stds
                        'Call EstablishCalStdGroups(False) 'qc stds
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try

                    'now redo arrAnalytes and redo tblanalysishome
                    'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                    '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                    '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET

                    'clear arranalytes
                    arrAnalytes.Clear(arrAnalytes, 0, arrAnalytes.Length)
                    ReDim arrAnalytes(intUBAA, 500)

                    var1 = UBound(arrAnalytes, 2) 'debug

                    Dim tblA As DataTable = tblAnalyteGroups
                    str1 = "Starting arrAnalytes..."
                    Dim intCTA As Short

                    Try
                        intCTA = 0
                        For Count1 = 0 To tblA.Rows.Count - 1
                            int1 = tblA.Rows(Count1).Item("INTGROUP")
                            strF = "INTGROUP = " & int1
                            'Dim rowsGG() As DataRow = tblCalStdGroupAssayIDsAcc.Select(strF)
                            Dim rowsGG() As DataRow = tblCalStdGroupAssayIDsAll.Select(strF)

                            If rowsGG.Length = 0 Then
                                var1 = var1
                            Else
                                intCTA = intCTA + 1
                                arrAnalytes(1, intCTA) = tblA.Rows(Count1).Item("ANALYTEDESCRIPTION_C")

                                arrAnalytes(2, intCTA) = rowsGG(0).Item("ANALYTEID")
                                arrAnalytes(3, intCTA) = rowsGG(0).Item("ANALYTEINDEX")

                                'get BQL, AQL, Conc Units tblallanalruns
                                'must first get a RunID from tbl
                                'tblCalStdGroupsAll
                                strF = "INTGROUP = " & int1
                                Dim rowCSGA() As DataRow = tblCalStdGroupsAll.Select(strF)
                                int2 = rowCSGA(0).Item("RUNID")
                                strF = "ANALYTEINDEX = " & rowsGG(0).Item("ANALYTEINDEX") & " AND ANALYTEID = " & rowsGG(0).Item("ANALYTEID") & " AND RUNID = " & int2
                                Dim rowAAR() As DataRow = tblAllAnalRuns.Select(strF)
                                'take first row
                                If rowAAR.Length = 0 Then
                                    arrAnalytes(4, intCTA) = "" 'BQL
                                    arrAnalytes(5, intCTA) = "" 'AQL
                                    arrAnalytes(6, intCTA) = "" 'Conc Units
                                Else
                                    arrAnalytes(4, intCTA) = rowAAR(0).Item("NM") 'BQL
                                    arrAnalytes(5, intCTA) = rowAAR(0).Item("VEC") 'AQL
                                    arrAnalytes(6, intCTA) = rowAAR(0).Item("CONCENTRATIONUNITS") 'Conc Units
                                End If


                                'find (7,
                                arrAnalytes(7, intCTA) = rowsGG.Length
                                arrAnalytes(8, intCTA) = "No"
                                arrAnalytes(9, intCTA) = "No"
                                var1 = NZ(tblA.Rows(Count1).Item("INTSTD"), "")
                                If Len(var1) = 0 Then
                                    arrAnalytes(10, intCTA) = "No"
                                Else
                                    arrAnalytes(10, intCTA) = "Yes"
                                End If
                                arrAnalytes(11, intCTA) = NZ(tblA.Rows(Count1).Item("INTSTD"), "")
                                arrAnalytes(12, intCTA) = rowsGG(0).Item("MASTERASSAYID")
                                arrAnalytes(13, intCTA) = False
                                arrAnalytes(14, intCTA) = tblA.Rows(Count1).Item("ANALYTEDESCRIPTION")
                                arrAnalytes(15, intCTA) = int1 'intgroup
                                var1 = tblA.Rows(Count1).Item("MATRIX")
                                arrAnalytes(16, intCTA) = var1 'tblA.Rows(Count1).Item("MATRIX")
                                arrAnalytes(17, intCTA) = intCTA

                                arrAnalytes(19, intCTA) = tblA.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
                                arrAnalytes(20, intCTA) = NZ(tblA.Rows(Count1).Item("INTSTD"), "")

                            End If

                        Next
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try


                    ctAnalytes = intCTA ' tblA.Rows.Count

                    ReDim arrAnalytesCB(ctAnalytes - 1)
                    For Count1 = 0 To ctAnalytes - 1
                        arrAnalytesCB(Count1) = arrAnalytes(1, Count1 + 1)
                    Next

                    ''debug
                    'var1 = arrAnalytes(16, 1)
                    'var1 = var1

                    'add back internal standards
                    ctAnalytes_IS = 0
                    int1 = 0

                    Count2 = ctAnalytes

                    Dim dvIS As DataView = New DataView(tblA, "", "INTSTD ASC", DataViewRowState.CurrentRows)
                    Dim tblIS As DataTable = dvIS.ToTable("a", True, "INTSTD")
                    Dim strMatrix As String

                    str1 = "Starting tblIS..."
                    ctAnalytes_IS = tblIS.Rows.Count
                    For Count1 = 1 To tblIS.Rows.Count

                        Count2 = Count2 + 1

                        var1 = tblIS.Rows(Count1 - 1).Item("INTSTD") 'debug
                        var2 = Replace(var1, ChrW(12288), " ", 1, -1, CompareMethod.Text)
                        arrAnalytes(1, Count2) = var2 ' rs.Fields("INTERNALSTDNAME").Value
                        'arrAnalytes(2, Count1) = rs.Fields("GlobalAnalyteID").Value 'GlobalAnalyteID=AnalyteID

                        'var1 = rs.Fields("ANALYTEINDEX") 'debug
                        'arrAnalytes(3, Count2) = rs.Fields("ANALYTEINDEX").Value 'NO!! IntStd has no analyteindex!!!


                        'Sheets("AnalRefTables").Range("AnalyteName").Offset(0, Count2).Value = arrAnalytes(1, Count2)
                        'Sheets("AnalRefTables").Range("IsReplicate").Offset(0, Count2).Value = "No"
                        arrAnalytes(8, Count2) = "No"
                        'Sheets("AnalRefTables").Range("IsInternalStandard").Offset(0, Count2).Value = "Yes"
                        arrAnalytes(9, Count2) = "Yes"
                        'Sheets("AnalRefTables").Range("UseInternalStandard").Offset(0, Count1).Value = "Yes"
                        arrAnalytes(10, Count2) = "NA"
                        'Sheets("AnalRefTables").Range("InternalStandard").Offset(0, Count1).Value = arrAnalytes(1, Count2)
                        arrAnalytes(11, Count2) = "NA"

                        'var1 = rs.Fields("MASTERASSAYID").Value 'debug
                        'arrAnalytes(12, Count2) = rs.Fields("MASTERASSAYID").Value 'NO!! IntStd has no masterassayid!!!

                        arrAnalytes(11, Count2) = NZ(var2, "")
                        arrAnalytes(13, Count2) = False ' is duplicate  :is coadministeredcmpd?
                        arrAnalytes(14, Count2) = var2
                        arrAnalytes(15, Count2) = -1
                        arrAnalytes(16, Count2) = ""
                        arrAnalytes(17, Count2) = Count2

                    Next Count1

                    ReDim Preserve arrAnalytes(intUBAA, ctAnalytes + ctAnalytes_IS)

                    'redo tbleAnalytesHome
                    tblAnalytesHome.Clear()

                    Dim boolFromTAG As Boolean

                    For Count1 = 1 To ctAnalytes + ctAnalytes_IS
                        drow8 = tblAnalytesHome.NewRow
                        For Count2 = 1 To intUBAA
                            boolFromTAG = False
                            str1 = ""
                            Select Case Count2
                                Case 1
                                    str1 = "AnalyteDescription" 'this is _C for multi-concentration
                                Case 2
                                    str1 = "AnalyteID"
                                Case 3
                                    str1 = "AnalyteIndex"
                                Case 4
                                    str1 = "BQL"
                                Case 5
                                    str1 = "AQL"
                                Case 6
                                    str1 = "ConcUnits"
                                Case 7
                                    str1 = "AcceptedRuns"
                                Case 8
                                    str1 = "IsReplicate"
                                Case 9
                                    str1 = "IsIntStd"
                                Case 10
                                    str1 = "UseIntStd"
                                Case 11
                                    str1 = "IntStd"
                                Case 12
                                    str1 = "MasterAssayID"
                                Case 13
                                    str1 = "IsCoadminCmpd"
                                Case 14
                                    str1 = "ORIGINALANALYTEDESCRIPTION"
                                Case 15
                                    str1 = "INTGROUP"
                                Case 16
                                    str1 = "MATRIX"
                                    boolFromTAG = True
                                Case 17
                                    str1 = "INTORDER"
                                Case 18
                                    str1 = "CALIBRSET"
                                    boolFromTAG = True

                                Case 19
                                    str1 = "CHARUSERANALYTE" '
                                Case 20
                                    str1 = "CHARUSERIS" '

                            End Select
                            If Len(str1) = 0 Then
                            Else
                                Try
                                    var2 = arrAnalytes(Count2, Count1) 'debug
                                    var1 = NZ(arrAnalytes(Count2, Count1), System.DBNull.Value)

                                    Select Case Count2
                                        Case 1
                                            strA = var1
                                    End Select
                                    If boolFromTAG Then
                                        strF = "ANALYTEDESCRIPTION_C = '" & CleanText(strA) & "'"
                                        Dim rowsTAG() As DataRow = tblAnalyteGroups.Select(strF)
                                        If rowsTAG.Length = 0 Then
                                        Else
                                            var1 = NZ(rowsTAG(0).Item(str1), System.DBNull.Value)
                                        End If
                                    End If

                                    Try
                                        drow8(str1) = var1
                                    Catch ex As Exception
                                        var1 = ex.Message
                                    End Try
                                Catch ex As Exception
                                    var1 = var1
                                End Try

                            End If
                        Next
                        Try
                            tblAnalytesHome.Rows.Add(drow8)
                        Catch ex As Exception
                            var1 = ex.Message
                        End Try

                    Next

                    'tblAnalytesHome.AcceptChanges()

                    'Call frmH.ConfigAnalyteOrder()

                End If
            Catch ex As Exception
                var1 = ex.Message
            End Try

            ''debug
            'var1 = arrAnalytes(16, 1)
            'var1 = var1

            'now compare tblAnalytesHome to tblStudyDocAnalytes
            'these two tables should be the same

            strF = "ID_TBLSTUDIES = " & id_tblStudies
            strS = "INTORDER ASC"

            ''check
            'int1 = TBLSTUDYDOCANALYTES.Columns.Count
            'int2 = tblAnalytesHome.Columns.Count



            Dim rowsSDA() As DataRow = TBLSTUDYDOCANALYTES.Select(strF, strS)

            Dim boolDoSDA As Boolean
            If rowsSDA.Length = 0 Then
                boolDoSDA = False
            Else
                boolDoSDA = True
            End If

            Dim intRowsTAH As Short = tblAnalytesHome.Rows.Count
            Dim intRowsSDA As Short = rowsSDA.Length

            '20180316 LEE:
            'Data may have changed within the study
            'TBLSTUDYDOCANALYTES must be updated accordingly
            'delete all entries in rowsSDA
            'delete SDA and start over

            'For Count2 = 0 To rowsSDA.Length - 1
            '    rowsSDA(Count2).Delete()
            'Next

            '20180321 LEE:
            'Hmm. Previous not smart, there's other tables that may need to be modified
            'Do this in a function
            Dim boolCSDA As Boolean = False

            boolCSDA = CheckStudyDocAnalytes1() 'False means OK

            ''******

            If rowsSDA.Length = 0 Or boolCSDA Then
                'record new entries
                Dim id As Int64
                id = GetMaxID("TBLSTUDYDOCANALYTES", tblAnalytesHome.Rows.Count, True)
                For Count2 = 0 To tblAnalytesHome.Rows.Count - 1

                    id = id + 1

                    Dim nr As DataRow = TBLSTUDYDOCANALYTES.NewRow

                    nr.BeginEdit()
                    nr.Item("ID_TBLSTUDYDOCANALYTES") = id
                    nr.Item("ID_TBLSTUDIES") = id_tblStudies
                    nr.Item("UPSIZE_TS") = dtNow

                    Try
                        For Count3 = 0 To tblAnalytesHome.Columns.Count - 1
                            str1 = tblAnalytesHome.Columns(Count3).ColumnName
                            var1 = tblAnalytesHome.Rows(Count2).Item(str1)
                            If Count3 = 15 Then
                                var1 = var1 'debug
                            End If
                            If TBLSTUDYDOCANALYTES.Columns.Contains(str1) Then
                                Try
                                    nr.Item(str1) = var1
                                Catch ex As Exception
                                    var2 = ex.Message
                                    var2 = var2
                                End Try
                            Else
                                var2 = var2
                            End If

                        Next
                    Catch ex As Exception
                        var2 = ex.Message
                        var2 = var2
                    End Try

                    nr.EndEdit()

                    TBLSTUDYDOCANALYTES.Rows.Add(nr)

                Next

                'Call PutMaxID("TBLSTUDYDOCANALYTES", id)

                'update datatable

                If boolGuWuOracle Then
                    Try
                        'ta_tblStudyDocAnalytes.Update(TBLSTUDYDOCANALYTES)
                    Catch ex As DBConcurrencyException

                    End Try

                ElseIf boolGuWuAccess Then
                    Try
                        ta_TBLSTUDYDOCANALYTESAcc.Update(TBLSTUDYDOCANALYTES)
                    Catch ex As DBConcurrencyException

                    End Try

                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_TBLSTUDYDOCANALYTESSQLSERVER.Update(TBLSTUDYDOCANALYTES)
                    Catch ex As DBConcurrencyException

                    End Try

                End If

                'now call CheckStudyDocAnalytes2
                If boolDoSDA Then
                    Call CheckStudyDocAnalytes2()
                End If



            Else


                Try

                    '20180316 LEE:
                    'something has happened such that some fields are not populated in TBLSTUDYDOCANALYTES
                    'must check to see what tblAnalytesHome has

                    '20180316 LEE:
                    'This logic is flawed. Data may have changed within the study, e.g. intOrder
                    'TBLSTUDYDOCANALYTES must be updated accordingly
                    '20180612 LEE:
                    'No!! tblAnalytesHome needs the upate and intOrder is the only item needing update

                    For Count2 = 0 To rowsSDA.Length - 1

                        var1 = rowsSDA(Count2).Item("AnalyteDescription")
                        int1 = rowsSDA(Count2).Item("IntOrder")
                        var2 = NZ(rowsSDA(Count2).Item("CHARUSERANALYTE"), "")
                        var3 = NZ(rowsSDA(Count2).Item("CHARUSERIS"), "")
                        strF = "AnalyteDescription = '" & CleanText(CStr(var1)) & "'" '20190206
                        Dim rowsTAH() As DataRow = tblAnalytesHome.Select(strF)
                        If rowsTAH.Length = 0 Then

                        Else
                            rowsTAH(0).BeginEdit()
                            rowsTAH(0).Item("intOrder") = int1
                            rowsTAH(0).Item("CHARUSERANALYTE") = var2
                            rowsTAH(0).Item("CHARUSERIS") = var3
                            rowsTAH(0).EndEdit()


                            'rowsSDA(Count2).BeginEdit()
                            'For Count3 = 0 To tblAnalytesHome.Columns.Count - 1
                            '    str1 = tblAnalytesHome.Columns(Count3).ColumnName

                            '    Try
                            '        var2 = rowsTAH(0).Item(str1)
                            '    Catch ex As Exception
                            '        var3 = ex.Message
                            '        var3 = var3
                            '    End Try

                            '    If Count3 = 15 Then
                            '        var1 = var1 'debug
                            '    End If

                            '    rowsSDA(Count2).Item(str1) = var2

                            'Next
                            'rowsSDA(0).EndEdit()
                        End If

                    Next Count2

                    ''20180316 LEE: depricate
                    'For Count2 = 0 To rowsSDA.Length - 1

                    '    var1 = rowsSDA(Count2).Item("AnalyteDescription")
                    '    strF = "AnalyteDescription = '" & var1 & "'"
                    '    Dim rowsTAH() As DataRow = tblAnalytesHome.Select(strF)
                    '    If rowsTAH.Length = 0 Then
                    '        'should never happen. ignore for now
                    '    Else
                    '        rowsTAH(0).BeginEdit()
                    '        For Count3 = 0 To tblAnalytesHome.Columns.Count - 1
                    '            str1 = tblAnalytesHome.Columns(Count3).ColumnName
                    '            If Count3 = 15 Then
                    '                var1 = var1 'debug
                    '            End If
                    '            Try
                    '                var2 = rowsSDA(Count2).Item(str1)
                    '            Catch ex As Exception
                    '                var3 = ex.Message
                    '                var3 = var3
                    '            End Try

                    '            Try
                    '                rowsTAH(0).Item(str1) = var2
                    '            Catch ex As Exception
                    '                var3 = ex.Message
                    '                var3 = var3
                    '            End Try

                    '        Next
                    '        rowsTAH(0).EndEdit()
                    '    End If

                    'Next Count2

                Catch ex As Exception
                    var3 = ex.Message
                    var3 = var3
                End Try

            End If

            tblAnalytesHome.AcceptChanges()

            ''debug
            'var1 = arrAnalytes(16, 1)
            'var1 = var1

            Call ReorderAnalytes()

            ''debug
            'var1 = arrAnalytes(16, 1)
            'var1 = var1

            Call frmH.ConfigAnalyteOrder()

            ''debug
            'For Count2 = 0 To tblAnalytesHome.Rows.Count - 1
            '    var1 = ""
            '    For Count3 = 0 To tblAnalytesHome.Columns.Count - 1
            '        var1 = var1 & ";" & tblAnalytesHome.Rows(Count2).Item(Count3)
            '    Next
            '    Console.WriteLine(var1)
            'Next


            '20150911: Larry End CalGroups

skipLarry1:


            Call ClearAllQCTables()

            str_cbxStudy = frmH.cbxStudy.Text

            'free up 2nd cols for data entry
            tblWatsonData.Columns.Item(1).ReadOnly = False

            str1 = "Retrieving Watson Data...1 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            Cursor.Current = Cursors.WaitCursor

            ''debug
            'var1 = arrAnalytes(16, 1)
            'var1 = var1

            'grab instance of CONFIGANALYTERUNSTATUS
            If rsRS.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsRS.Close()
            End If
            rsRS.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            If boolAccess Then
                str1 = "SELECT * FROM CONFIGANALYTERUNSTATUS;"
            Else
                str1 = "SELECT * FROM " & strSchema & ".CONFIGANALYTERUNSTATUS;"
            End If

            rsRS.Open(str1, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsRS.ActiveConnection = Nothing

            If frmH.dgvwStudy.CurrentRow Is Nothing Then
                'probably comes from new oracle record
                var1 = wStudyID
                var2 = wProjectID
                var3 = wWStudyName
            End If

            var1 = frmH.dgvwStudy.Rows.Count
            int1 = frmH.dgvwStudy.CurrentRow.Index

            '20190130 LEE:
            'For some reason, tblWS is retrieving unfiltered dgvwStudy.datasource
            'instead, use dgvwStudy directly
            'Dim tblWS As DataTable
            'tblWS = frmH.dgvwStudy.DataSource
            'var1 = tblWS.Rows.Count 'debug
            'wStudyID = tblWS.Rows(int1).Item("STUDYID")
            'wProjectID = tblWS.Rows(int1).Item("PROJECTID")
            'wWStudyName = tblWS.Rows(int1).Item("STUDYNAME")

            wStudyID = frmH.dgvwStudy("STUDYID", int1).Value
            wProjectID = frmH.dgvwStudy("PROJECTID", int1).Value
            wWStudyName = frmH.dgvwStudy("STUDYNAME", int1).Value

            ctRows = 0

            With tblWatsonData
                int1 = FindRow("Watson Study ID", tblWatsonData, "Item")
                .Rows.Item(int1).Item(1) = wStudyID
            End With

            With tblWatsonData
                int1 = FindRow("Watson Project ID", tblWatsonData, "Item")
                .Rows.Item(int1).Item(1) = wProjectID
            End With

            str1 = "Retrieving Watson Data...2 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'record speciesid from a different recordset
            If boolAccess Then
                str1 = "SELECT Assay.* "
                str2 = "FROM Assay "
                str3 = "WHERE StudyID = " & wStudyID & ";"
            Else
                str1 = "SELECT " & strSchema & ".Assay.* "
                str2 = "FROM " & strSchema & ".Assay "
                str3 = "WHERE " & strSchema & ".Assay.StudyID = " & wStudyID & ";"
            End If

            strSQL = str1 & str2 & str3
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing
            var1 = ""
            Do Until rs.EOF
                var1 = NZ(rs.Fields("SpeciesID").Value, "")
                If Len(var1) > 0 Then
                    Exit Do
                End If
                rs.MoveNext()
            Loop
            rs.Close()
            If Len(NZ(var1, "")) = 0 Then
                MsgBox("Hmmm. The study species is not configured in the Watson database. Please investigate and correct. This Workbook Preparation action is terminated.", vbInformation + vbOKOnly, "Species must be configured...")
                wSpeciesID = 0
                GoTo end1
            End If
            wSpeciesID = CLng(var1)

            'With tblWatsonData
            '    int1 = FindRow("Species", tblWatsonData)
            '    .Rows.item(int1).Item(1) = wSpeciesID
            'End With

            str1 = "Retrieving Watson Data...3 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'new table to straighten out difficult CalibrStd NomConc
            If boolAccess Then
                str1 = "SELECT ASSAYANALYTEKNOWN.ASSAYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTES.NM, ASSAYANALYTES.VEC "
                str2 = "FROM ASSAYANALYTES INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE((  (((ASSAYANALYTEKNOWN.KNOWNTYPE) = 'STANDARD') Or ((ASSAYANALYTEKNOWN.KNOWNTYPE) = 'QC')))  And ((ASSAYANALYTEKNOWN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTEKNOWN.ASSAYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            Else
                str1 = "SELECT " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE(((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) = 'STANDARD') And ((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"
            End If

            strSQL = str1 & str2 & str3 & str4
            '''Console.WriteLine(strSQL)

            Dim rsAAUnk As New ADODB.Recordset
            rsAAUnk.CursorLocation = CursorLocationEnum.adUseClient
            rsAAUnk.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsAAUnk.ActiveConnection = Nothing

            tblAAUnk.Clear()
            tblAAUnk.AcceptChanges()
            tblAAUnk.BeginLoadData()
            daDoPr.Fill(tblAAUnk, rsAAUnk)
            tblAAUnk.EndLoadData()


            'get reassay information

            Dim str5a As String
            If boolAccess Then
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNTREATMENT.TREATMENTID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNTREATMENT.TREATMENTKEY, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE "
                str2 = "FROM DESIGNTREATMENT INNER JOIN (((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN SAMPLERESULTSCONFLICT ON (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) ON (DESIGNTREATMENT.TREATMENTKEY = DESIGNSAMPLE.TREATMENTEVENTID) AND (DESIGNTREATMENT.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                'this statement adds reasons
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNTREATMENT.TREATMENTID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNTREATMENT.TREATMENTKEY, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE "
                str2 = "FROM SAMPRESCONFLICTDEC INNER JOIN (DESIGNTREATMENT INNER JOIN (((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN SAMPLERESULTSCONFLICT ON (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID)) ON (DESIGNTREATMENT.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNTREATMENT.TREATMENTKEY = DESIGNSAMPLE.TREATMENTEVENTID)) ON (SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPRESCONFLICTDEC.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) "
                'str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ")) "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                'NOTE: don't know why DESIGNTREATMENT is in here. in example XLA10BX, this table is almost blank, which
                '      causes recordset to be empty
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE "
                '20150703 Larry: remove DecisionCode - Never used and adds unecessary rows to query
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE "
                str2 = "FROM (((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN SAMPLERESULTSCONFLICT ON (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN SAMPRESCONFLICTDEC ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTDEC.STUDYID) "
                str3 = "WHERE(((SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ") And ((SAMPRESCONFLICTDEC.REASSAYREASON) <> 'NA')) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20150704 Larry: Redo of sql, using all primary keys, no DecisionCode
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE "
                str2 = "FROM (((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN SAMPLERESULTSCONFLICT ON (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN SAMPRESCONFLICTDEC ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTDEC.STUDYID) "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                'NDL 28-Jan-2016  The original query wasn't taking into account the case where two decisions were made on the same sample
                '(presumably because the user changed their mind).  This fixes it and clears the query up a bit.
                'NDL 7-Feb-2016 Added info for rsTimePpoint in DoPrepare: DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM (SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID)) ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160220 LEE: Added WEEK
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM ((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160531 LEE: Need to filter DesignSubjectID > 0
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM ((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((DESIGNSUBJECT.DESIGNSUBJECTID)>0)) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20180124 LEE: Added matrix SAMPLETYPEID
                '                SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID
                'FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID)) ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY
                'WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=1322) AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((DESIGNSUBJECT.DESIGNSUBJECTID)>0))
                'ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;

                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID)) ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((DESIGNSUBJECT.DESIGNSUBJECTID)>0)) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20180227 LEE:
                'need to add , DESIGNSAMPLE.USERSAMPLEID
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID, DESIGNSAMPLE.USERSAMPLEID "
                str2 = "FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID)) ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((DESIGNSUBJECT.DESIGNSUBJECTID)>0)) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"
                'USERSAMPLEID
            Else

                If boolANSI Then
                    str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNTREATMENT.TREATMENTID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNTREATMENT.TREATMENTKEY, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE "
                    str2 = "FROM DESIGNTREATMENT INNER JOIN (((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN SAMPLERESULTSCONFLICT ON (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) ON (DESIGNTREATMENT.TREATMENTKEY = DESIGNSAMPLE.TREATMENTEVENTID) AND (DESIGNTREATMENT.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) "
                    str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ")) "
                    str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    'this statement adds reasons
                    str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNTREATMENT.TREATMENTID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNTREATMENT.TREATMENTKEY, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE "
                    str2 = "FROM SAMPRESCONFLICTDEC INNER JOIN (DESIGNTREATMENT INNER JOIN (((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN SAMPLERESULTSCONFLICT ON (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID)) ON (DESIGNTREATMENT.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNTREATMENT.TREATMENTKEY = DESIGNSAMPLE.TREATMENTEVENTID)) ON (SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPRESCONFLICTDEC.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) "
                    'str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ")) "
                    str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                    str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    'NOTE: don't know why DESIGNTREATMENT is in here. in example XLA10BX, this table is almost blank, which
                    '      causes recordset to be empty
                    str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE "
                    '20150703 Larry: remove DecisionCode - Never used and adds unecessary rows to query
                    str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE "
                    str2 = "FROM (((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".SAMPLERESULTSCONFLICT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) "
                    str3 = "WHERE(((" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON) <> 'NA')) "
                    str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    ''20150704 Larry: Redo of sql, using all primary keys, no DecisionCode
                    str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE "
                    str2 = "FROM (((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".SAMPLERESULTSCONFLICT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) "
                    str3 = "WHERE (((" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                    str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY," & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    'NDL 28-Jan-2016  The original query wasn't taking into account the case where two decisions were made on the same sample
                    '(presumably because the user changed their mind).  This fixes it and clears the query up a bit.
                    str1 = "SELECT DISTINCT  " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID,  " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG,  " _
                        & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME,  " & strSchema & ".DESIGNSAMPLE.ENDDAY,  " _
                        & strSchema & ".DESIGNSAMPLE.ENDHOUR,  " & strSchema & ".DESIGNSAMPLE.ENDMINUTE,  " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID,  " _
                        & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER,  " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID,  " _
                        & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID,  " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID,  " _
                        & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON,  " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON,  " _
                        & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE,  " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE,  " _
                        & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP,  " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP,  " _
                        & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP,  " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP "
                    str2 = "FROM (" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" _
                        & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" _
                        & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (" _
                        & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" _
                        & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" _
                        & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)) ON (" _
                        & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (" _
                        & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" _
                        & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID)) ON (" _
                        & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (" _
                        & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" _
                        & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" _
                        & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((" _
                        & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" _
                        & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" _
                        & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID)) INNER JOIN " _
                        & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " _
                        & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " _
                        & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " _
                        & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " _
                        & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) "
                    str3 = "WHERE ((( " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND (( " _
                        & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ")) "
                    str4 = "ORDER BY  " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG,  " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME,  " _
                        & strSchema & ".DESIGNSAMPLE.ENDDAY,  " & strSchema & ".DESIGNSAMPLE.ENDHOUR,  " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID,  " _
                        & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    '20160220 LEE: Added WEEK 
                    str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND," & strSchema & ". DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND "
                    str2 = "FROM ((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (S" & strSchema & ".AMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                    str3 = "WHERE (((" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                    str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    '20160531 LEE: Need to filter DesignSubjectID > 0
                    str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND," & strSchema & ". DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND "
                    str2 = "FROM ((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                    str3 = "WHERE (((" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)>0)) "
                    str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"


                    '20180124 LEE: Added matrix SAMPLETYPEID
                    '                SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND, CONFIGSAMPLETYPES.SAMPLETYPEID
                    'FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID)) ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY
                    'WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=1322) AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((DESIGNSUBJECT.DESIGNSUBJECTID)>0))
                    'ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;
                    str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND," & strSchema & ". DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                    'str2 = "FROM ((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                    str2 = "FROM (((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)>0)) "
                    str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"


                    '20180227 LEE:
                    'need to add , DESIGNSAMPLE.USERSAMPLEID
                    str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND," & strSchema & ". DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID "
                    str2 = "FROM (((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) ON (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)>0)) "
                    str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"
                    'USERSAMPLEID


                Else
                    'str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNTREATMENT.TREATMENTID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNTREATMENT.TREATMENTKEY, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE "
                    'str2 = "FROM DESIGNTREATMENT, DESIGNSAMPLE, DESIGNSUBJECT, DESIGNSUBJECTGROUP, SAMPLERESULTSCONFLICT "
                    'str3 = "WHERE ((((DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID)) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) AND (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) AND (DESIGNTREATMENT.TREATMENTKEY = DESIGNSAMPLE.TREATMENTEVENTID) AND (DESIGNTREATMENT.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) "
                    'str3 = str3 & "AND (((SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & ")) "
                    'str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    ''this statement adds reasons
                    'str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNTREATMENT.TREATMENTID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNTREATMENT.TREATMENTKEY, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE "
                    'str2 = "FROM SAMPRESCONFLICTDEC, DESIGNTREATMENT, DESIGNSAMPLE, DESIGNSUBJECT, DESIGNSUBJECTGROUP, SAMPLERESULTSCONFLICT "
                    'str3 = "WHERE (((((DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) AND (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID)) AND (DESIGNTREATMENT.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (DESIGNTREATMENT.TREATMENTKEY = DESIGNSAMPLE.TREATMENTEVENTID)) AND (SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPRESCONFLICTDEC.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) "
                    'str3 = str3 & "AND (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                    'str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    ''NOTE: don't know why DESIGNTREATMENT is in here. in example XLA10BX, this table is almost blank, which
                    ''      causes recordset to be empty
                    'str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE"
                    'str2 = "FROM SAMPRESCONFLICTDEC, DESIGNSAMPLE, DESIGNSUBJECT, DESIGNSUBJECTGROUP, SAMPLERESULTSCONFLICT "
                    'str3 = "WHERE (((((DESIGNSAMPLE.SUBJECTGROUPID = DESIGNSUBJECT.SUBJECTGROUPID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) AND (DESIGNSAMPLE.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (DESIGNSAMPLE.STUDYID = SAMPLERESULTSCONFLICT.STUDYID))) AND (SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPRESCONFLICTDEC.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) "
                    'str3 = str3 & "AND (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                    'str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                    ''NDL: Don't know what "boolANSI" means - 28-jan-2016 - Larry please look at this.
                    ''20160220 LEE: Oracle 8.n and ealier was not ANSI (American National Standards Institute) compliant, so SQL statements were drastically different from SQL Server, Sybase, MySQL, etc., whose languages are ANSI compliant
                    ''starting with Oracle 11.n, Oracle became ANSI compliant. DoPrepare still has pre-Oracle 11 (boolANSI = False) code that can be ignored and eventually removed
                    'str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.DESIGNSAMPLEID, SAMPLERESULTSCONFLICT.STUDYID, DESIGNSAMPLE.TREATMENTEVENTID, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP "
                    'str2 = "FROM (SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID)) ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) "
                    'str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID)=" & wStudyID & ") AND ((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA')) "
                    'str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                End If
            End If


            strSQL = str1 & str2 & str3 & str4

            'Console.WriteLine("tblReassayReport: " & strSQL)

            'str5a = str1 & str2 & str3 & str4 'for debugging

            Count2 = 0
            Dim rsReassay1 As New ADODB.Recordset

            rsReassay1.CursorLocation = CursorLocationEnum.adUseClient
            rsReassay1.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsReassay1.ActiveConnection = Nothing

            tblReassayReport.Clear()
            tblReassayReport.AcceptChanges()
            tblReassayReport.BeginLoadData()
            daDoPr.Fill(tblReassayReport, rsReassay1)
            tblReassayReport.EndLoadData()

            rsReassay1.Close()
            rsReassay1 = Nothing

            'open and store reassay reasons recordset
            If boolAccess Then
                str1 = "SELECT SAMPRESCONFLICTDEC.* "
                str2 = "FROM SAMPRESCONFLICTDEC "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.STUDYID) = " & wStudyID & "));"
            Else
                str1 = "SELECT " & strSchema & ".SAMPRESCONFLICTDEC.* "
                str2 = "FROM " & strSchema & ".SAMPRESCONFLICTDEC "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) = " & wStudyID & "));"
            End If

            strSQL = str1 & str2 & str3

            ''''''''''''''''''Console.WriteLine("rsReassayReason:tblReassayReasons")
            ''''Console.WriteLine("rsReassayReason: " & strSQL)
            Dim rsReassayReason As New ADODB.Recordset
            rsReassayReason.CursorLocation = CursorLocationEnum.adUseClient
            rsReassayReason.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsReassayReason.ActiveConnection = Nothing

            'save this recordset in a datatable
            'add columns to table tblReassay
            tblReassayReasons.Clear()
            tblReassayReasons.AcceptChanges()
            tblReassayReasons.BeginLoadData()
            daDoPr.Fill(tblReassayReasons, rsReassayReason)
            tblReassayReasons.EndLoadData()

            rsReassayReason.Close()
            rsReassayReason = Nothing

            'open and store reassayResults recordset
            If boolAccess Then
                str1 = "SELECT SAMPLERESULTSCONFLICT.* "
                str2 = "FROM SAMPLERESULTSCONFLICT "
                str3 = "WHERE (((SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & "));"
            Else
                str1 = "SELECT " & strSchema & ".SAMPLERESULTSCONFLICT.* "
                str2 = "FROM " & strSchema & ".SAMPLERESULTSCONFLICT "
                str3 = "WHERE (((" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) = " & wStudyID & "));"
            End If

            strSQL = str1 & str2 & str3

            ''''''''''''''''''Console.WriteLine("rsSAMPLERESULTSCONFLICT:rsSAMPLERESULTSCONFLICT")
            ''''''''''''''''''Console.WriteLine(strSQL)
            Dim rsSAMPLERESULTSCONFLICT As New ADODB.Recordset

            rsSAMPLERESULTSCONFLICT.CursorLocation = CursorLocationEnum.adUseClient
            rsSAMPLERESULTSCONFLICT.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsSAMPLERESULTSCONFLICT.ActiveConnection = Nothing

            'save this recordset in a datatable
            'add columns to table tblReassay
            tblSAMPLERESULTSCONFLICT.Clear()
            tblSAMPLERESULTSCONFLICT.AcceptChanges()
            tblSAMPLERESULTSCONFLICT.BeginLoadData()
            daDoPr.Fill(tblSAMPLERESULTSCONFLICT, rsSAMPLERESULTSCONFLICT)
            tblSAMPLERESULTSCONFLICT.EndLoadData()

            rsSAMPLERESULTSCONFLICT.Close()
            rsSAMPLERESULTSCONFLICT = Nothing



            'enter WAtson information in tblWatson
            '20171106 LEE: Matrix may be repeated because user hasn't recorded a STANDARDVOLUME
            Dim dvXX As DataView = New DataView(tblSpeciesMatrix)
            Dim tblXX As DataTable = dvXX.ToTable("XX", True, "SAMPLETYPEID")
            Try
                For Count1 = 1 To tblXX.Rows.Count
                    str1 = NZ(tblXX.Rows(Count1 - 1).Item("SAMPLETYPEID"), "NA")
                    If Count1 = 1 Then
                        str2 = str1
                    ElseIf Count1 = tblXX.Rows.Count Then
                        If Count1 = 2 Then
                            str2 = str2 & " and " & str1
                        Else
                            str2 = str2 & ", and " & str1
                        End If

                    Else
                        str2 = str2 & ", " & str1
                    End If
                Next
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try


            With tblWatsonData
                int1 = FindRow("Matrix", tblWatsonData, "Item")
                .Rows.Item(int1).Item(1) = str2
            End With

            'create a STANDARDVOLUME
            Try
                int1 = 0
                str2 = "NA"
                For Count1 = 1 To tblXX.Rows.Count
                    str1 = NZ(tblXX.Rows(Count1 - 1).Item("SAMPLETYPEID"), "NA")
                    Dim rowsSV() As DataRow = tblSpeciesMatrixSV.Select("SAMPLETYPEID = '" & str1 & "'")
                    If rowsSV.Length = 0 Then
                    Else
                        int1 = int1 + 1
                        If rowsSV.Length > 1 Then
                            For Count2 = 0 To rowsSV.Length - 1
                                var1 = rowsSV(Count2).Item("STANDARDVOLUME")
                                If IsDBNull(var1) Then
                                    str3 = "0"
                                Else
                                    str3 = CStr(var1)
                                    Exit For
                                End If
                            Next

                        Else
                            str3 = NZ(rowsSV(0).Item("STANDARDVOLUME"), "0")
                        End If

                        If int1 = 1 Then
                            str2 = str3
                        ElseIf Count1 = tblXX.Rows.Count Then
                            If int1 = 2 Then
                                str2 = str2 & " and " & str3
                            Else
                                str2 = str2 & ", and " & str3
                            End If

                        Else
                            str2 = str2 & ", " & str3
                        End If
                    End If

                Next
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try


            ''debug
            'var1 = arrAnalytes(16, 1)
            'var1 = var1

            'var1 = NZ(tblSpeciesMatrixSV.Rows(0).Item("STANDARDVOLUME"), 0)
            'With tblWatsonData
            '    int1 = FindRow("Sample Size", tblWatsonData, "Item")
            '    .Rows.Item(int1).Item(1) = var1
            'End With

            With tblWatsonData
                int1 = FindRow("Sample Size", tblWatsonData, "Item")
                .Rows.Item(int1).Item(1) = str2
            End With

            var1 = NZ(tblSpeciesMatrix.Rows(0).Item("SPECIES"), 0)
            With tblWatsonData
                int1 = FindRow("Species", tblWatsonData, "Item")
                .Rows.Item(int1).Item(1) = var1
            End With
            'Sheets("Data").Range("Species").Offset(0, 1).Value = UnCapit(CStr(var1))

            'do Integration Type after tblRegCon

            'bind data to datagrid
            'dgDataWatson.DataSource = tblWatsonData
            tblWatsonData.Columns.Item(1).ReadOnly = True
            frmH.dgvDataWatson.Refresh()
            frmH.dgvDataWatson.Columns(0).ReadOnly = True
            frmH.dgvDataWatson.Columns(1).ReadOnly = True
            frmH.dgvDataWatson.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            frmH.dgvDataWatson.AutoResizeColumns()
            frmH.dgvDataWatson.AutoResizeRows()

            System.Windows.Forms.Application.DoEvents()

            str1 = "Retrieving Watson Data...4 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            '***Start here 2


            str1 = "Retrieving Watson Data...6 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            If ctPB > frmH.pb1.Maximum Then
                ctPB = 1
            End If
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()


            'now do only analytes. Previous code produces replicates

            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, GLOBALANALYTES.PROJECTID "
                str2 = "FROM ANARUNANALYTERESULTS INNER JOIN (ASSAY INNER JOIN ((ASSAYANALYTES INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID) INNER JOIN STUDY ON ASSAYANALYTES.STUDYID = STUDY.STUDYID) ON (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) ON (ANARUNANALYTERESULTS.STUDYID = ASSAY.STUDYID) AND (ANARUNANALYTERESULTS.RUNID = ASSAY.RUNID) "
                str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & wStudyID & ") And ((GLOBALANALYTES.ACTIVE) = -1) And ((ASSAY.RUNID) > 0)) "
                str4 = "ORDER BY GLOBALANALYTES.ANALYTEDESCRIPTION;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".GLOBALANALYTES.PROJECTID "
                str2 = "FROM " & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (" & strSchema & ".ASSAY INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN " & strSchema & ".GLOBALANALYTES ON " & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID) INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".STUDY.STUDYID) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) ON (" & strSchema & ".ANARUNANALYTERESULTS.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ASSAY.RUNID) "
                str3 = "WHERE(((" & strSchema & ".ASSAYANALYTES.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".GLOBALANALYTES.ACTIVE) = -1) And ((" & strSchema & ".ASSAY.RUNID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION;"
            End If

            strSQL = str1 & str2 & str3 & str4

            Dim rsA As New ADODB.Recordset
            rsA.CursorLocation = CursorLocationEnum.adUseClient
            If rsA.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsA.Close()
            End If
            rsA.CursorLocation = CursorLocationEnum.adUseClient
            rsA.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsA.ActiveConnection = Nothing

            tblAnalU.Clear()
            tblAnalU.AcceptChanges()
            tblAnalU.BeginLoadData()
            daDoPr.Fill(tblAnalU, rsA)
            tblAnalU.EndLoadData()

            If rsA.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsA.Close()
            End If
            rsA = Nothing

            'after this, need to find this info for calibr and QC stds
            'for each analyte
            '  - each MasterAssy
            'for each Assay within each MasterAssay
            '  - Calibr NomConc
            '  - QC NomConc
            '  - Now need to compare NomConcs in each Assay to deterime unique sets of Calibr and QC NomConc's



            '***Start here 3
            str1 = "Retrieving calibration standard BLQ and ALQ info..."
            'frm.frmh.lblProgress.Text = str1
            'frm.Refresh()

            'determine BQL,AQL, concentration units
            If boolUseGroups Then
                'assigned earlier
            Else


                '20150911 Larry: these aren't used, deprecate it soon
                '20160205 LEE: Not true. These are used, but still need to deprecate

                For Count1 = 1 To ctAnalytes

                    '20160205 
                    If boolAccess Then
                        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAY.MASTERASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, CONCENTRATIONUNITS.CONCENTRATIONUNITS "
                        str2 = "FROM ((ASSAY INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAY.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) INNER JOIN ASSAYANALYTES ON ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) INNER JOIN CONCENTRATIONUNITS ON ASSAYANALYTES.CONCUNITSID = CONCENTRATIONUNITS.CONCUNITSID "
                        str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANARUNANALYTERESULTS.ANALYTEINDEX)=" & arrAnalytes(3, Count1) & ") AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((ASSAY.MASTERASSAYID)=" & arrAnalytes(12, Count1) & ") AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3));"
                    Else
                        If boolANSI Then
                            str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC, " & strSchema & ".CONCENTRATIONUNITS.CONCENTRATIONUNITS "
                            str2 = "FROM ((" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON " & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) INNER JOIN " & strSchema & ".CONCENTRATIONUNITS ON " & strSchema & ".ASSAYANALYTES.CONCUNITSID = " & strSchema & ".CONCENTRATIONUNITS.CONCUNITSID "
                            str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)=" & arrAnalytes(3, Count1) & ") AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((" & strSchema & ".ASSAY.MASTERASSAYID)=" & arrAnalytes(12, Count1) & ") AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3));"
                        Else
                            str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAY.MASTERASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, CONCENTRATIONUNITS.CONCENTRATIONUNITS "
                            str2 = "FROM ASSAY, ANALYTICALRUNANALYTES, ANALYTICALRUNSAMPLE, ANARUNANALYTERESULTS, ASSAYANALYTES, CONCENTRATIONUNITS "
                            str2 = str2 & "WHERE (((((ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) AND (ASSAY.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) AND ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND ASSAYANALYTES.CONCUNITSID = CONCENTRATIONUNITS.CONCUNITSID "
                            str3 = "AND (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANARUNANALYTERESULTS.ANALYTEINDEX)=" & arrAnalytes(3, Count1) & ") AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((ASSAY.MASTERASSAYID)=" & arrAnalytes(12, Count1) & ") AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3));"
                        End If
                    End If

                    strSQL = str1 & str2 & str3
                    ''Console.WriteLine(strSQL)
                    Try
                        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                            rs.Close()
                        End If
                        rs.CursorLocation = CursorLocationEnum.adUseClient
                        rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        rs.ActiveConnection = Nothing
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try



                    Try
                        If rs.RecordCount = 0 Then
                            arrAnalytes(4, Count1) = 0 ' rs.Fields("NM").Value
                            arrAnalytes(5, Count1) = 0 ' rs.Fields("VEC").Value
                            arrAnalytes(6, Count1) = 0 ' rs.Fields("CONCENTRATIONUNITS").Value

                        Else
                            arrAnalytes(4, Count1) = rs.Fields("NM").Value
                            arrAnalytes(5, Count1) = rs.Fields("VEC").Value
                            arrAnalytes(6, Count1) = rs.Fields("CONCENTRATIONUNITS").Value

                        End If
                    Catch ex As Exception
                        var1 = ex.Message
                    End Try

                    If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rs.Close()
                    End If
                Next
            End If

            str1 = "Retrieving Watson Data...7 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            '***End here 3

            ''fill tblWatsonAnalRefTable after 
            Call AddColumnsWatsonAnalRefTable()
            Call FillWatsonAnalRefTable()
            'dgWatsonAnalRef.Refresh()

            '***Start here 4

            'find number of Analytical Runs from table AnalyticalRunAnalytes
            str1 = "Retrieving analytical run info..."
            'frm.frmh.lblProgress.Text = str1
            'frm.Refresh()
            'frmh.lblProgress.Text = str1
            'frmh.lblProgress.Refresh()

            If boolANSI Then
                str1 = "SELECT ANALYTICALRUNANALYTES.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNANALYTES.ACCEPTREJECTREASON, ANALYTICALRUN.RUNDESCRIPTION, ANALYTICALRUN.NOTEBOOK, ANALYTICALRUN.PAGENUMBER, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE, ANALYTICALRUNANALYTES.STUDYID, ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
                str3 = "WHERE (((ANALYTICALRUNANALYTES.STUDYID)= " & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUNANALYTES.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

                'ALL ANALYTICAL RUNS!!!!
                'the following SQL returns all configured analytical runs. May convert to this later
                'str1 = "SELECT ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNANALYTES.ACCEPTREJECTREASON, ANALYTICALRUN.RUNDESCRIPTION, ANALYTICALRUN.NOTEBOOK, ANALYTICALRUN.PAGENUMBER, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE, ANALYTICALRUN.STUDYID, ANALYTICALRUN.RUNTYPEID "
                'str2 = "FROM ANALYTICALRUNANALYTES RIGHT JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
                'str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                'str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"

            Else
                str1 = "SELECT ANALYTICALRUNANALYTES.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNANALYTES.ACCEPTREJECTREASON, ANALYTICALRUN.RUNDESCRIPTION, ANALYTICALRUN.NOTEBOOK, ANALYTICALRUN.PAGENUMBER, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE, ANALYTICALRUNANALYTES.STUDYID, ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ANALYTICALRUNANALYTES, ANALYTICALRUN "
                str2 = str2 & "WHERE (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
                str3 = "AND (((ANALYTICALRUNANALYTES.STUDYID)= " & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUNANALYTES.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"
            End If

            'This returns ALL runs, but results in null for the following columns:
            'ANALYTEINDEX
            'RUNANALYTEREGRESSIONSTATUS
            'ACCEPTREJECTREASON
            'RUNSTARTED
            'EXTRACTIONDATE
            str1 = "SELECT ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNANALYTES.ACCEPTREJECTREASON, ANALYTICALRUN.RUNDESCRIPTION, ANALYTICALRUN.NOTEBOOK, ANALYTICALRUN.PAGENUMBER, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUN.EXTRACTIONDATE, ANALYTICALRUN.STUDYID, ANALYTICALRUN.RUNTYPEID "
            str2 = "FROM ANALYTICALRUNANALYTES RIGHT JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
            str4 = "ORDER BY ANALYTICALRUN.RUNID, ANALYTICALRUNANALYTES.ANALYTEINDEX;"
            strSQL = str1 & str2 & str3 & str4

            'REDO THIS THINKING
            'need to record every analytical run


            If boolAccess Then
                str1 = "SELECT ANALYTICALRUN.* "
                str2 = "FROM ANALYTICALRUN "
                str3 = "WHERE (((ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID;"
            Else
                str1 = "SELECT " & strSchema & ".ANALYTICALRUN.* "
                str2 = "FROM " & strSchema & ".ANALYTICALRUN "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID;"

            End If

            strSQL = str1 & str2 & str3 & str4


            'Console.WriteLine("tblAnalRunSum: " & strSQL)

            'set cursor as local because will filter later
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            ctAnalyticalRuns = rs.RecordCount ' / ctAnalytes

            ReDim arrAnalyticalRuns(15, ctAnalyticalRuns * ctAnalytes)
            '1=RUNID, 2=NOTEBOOK-PAGENUMBER, 3=EXTRACTIONDATE, 4=RUNSTARTDATE, 5=ANAREGSTATUSDESC
            '6=RUNDESCRIPTION, 7=ACCEPTREJECTREASON, 8=ANALYTE, 9=RUNTYPEID, 10=NM, 11=VEC, 12=ANALYTEID, 13=BOOLINTHISRUNSASSAYID, 14=RUNANALYTEREGRESSIONSTATUS,15=INSTGROUPNAME

            If boolAccess Then
                str1 = "SELECT ANALYTICALRUNANALYTES.* "
                str2 = "FROM ANALYTICALRUNANALYTES "
                str3 = "WHERE (((ANALYTICALRUNANALYTES.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY  ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNID;"

                '20171108 LEE: Alturas wants to add INSTGROUPNAME
                str1 = "SELECT ANALYTICALRUNANALYTES.*, ANALYTICALRUN.INSTGROUPNAME "
                str2 = "FROM ANALYTICALRUN RIGHT JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) "
                str3 = "WHERE (((ANALYTICALRUNANALYTES.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY  ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNID;"


            Else
                str1 = "SELECT " & strSchema & ".ANALYTICALRUNANALYTES.* "
                str2 = "FROM " & strSchema & ".ANALYTICALRUNANALYTES "

                '20171108 LEE: Alturas wants to add INSTGROUPNAME
                str1 = "SELECT " & strSchema & ".ANALYTICALRUNANALYTES.*, " & strSchema & ".ANALYTICALRUN.INSTGROUPNAME "
                str2 = "FROM " & strSchema & ".ANALYTICALRUN RIGHT JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY  " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNID;"
            End If
            strSQL = str1 & str2 & str3 & str4

            'Console.WriteLine("tblAnalRunSum: " & strSQL)

            'set cursor as local because will filter later
            If rs1.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs1.Close()
            End If
            rs1.CursorLocation = CursorLocationEnum.adUseClient
            rs1.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs1.ActiveConnection = Nothing

            Count1 = 0
            Dim arrRunDates() ' As Date
            Dim arrExtDates() ' As Date

            numQCLevels = 0
            numRepDilnQC = 0
            numRepQC = 0

            Dim intAAA As Int16

            For Count2 = 1 To ctAnalytes

                'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET
                var1 = arrAnalytes(14, Count2) 'debug

                If rs.EOF And rs.BOF Then
                Else
                    rs.MoveFirst()
                End If

                For Count3 = 1 To ctAnalyticalRuns

                    int1 = rs.Fields("RUNID").Value

                    Count1 = Count1 + 1

                    strF = "ANALYTEINDEX = " & arrAnalytes(3, Count2) & " AND RUNID = " & int1 'keep this. no reference to analyteid
                    rs1.Filter = ""
                    rs1.Filter = strF

                    intAAA = rs1.RecordCount 'debug

                    If rs1.EOF Then

                        '1=RUNID, 2=NOTEBOOK-PAGENUMBER, 3=EXTRACTIONDATE, 4=RUNSTARTDATE, 5=ANAREGSTATUSDESC
                        '6=RUNDESCRIPTION, 7=ACCEPTREJECTREASON, 8=ANALYTE, 9=RUNTYPEID, 10=NM, 11=VEC, 12=ANALYTEID, 13=BOOLINTHISRUNSASSAYID, 
                        '14=RUNANALYTEREGRESSIONSTATUS,15=INSTGROUPNAME


                        arrAnalyticalRuns(1, Count1) = int1 'rs.Fields("RUNID").Value
                        'evalutate pagenumber
                        var1 = NZ(rs.Fields("NOTEBOOK").Value, "")
                        var2 = NZ(rs.Fields("PAGENUMBER").Value, "")
                        If InStr(1, var2, var1, vbTextCompare) > 0 Then
                            var3 = var2
                        ElseIf Len(var2) = 0 Then
                            var3 = var1
                        Else
                            var3 = var1 & "-" & var2
                        End If
                        arrAnalyticalRuns(2, Count1) = NZ(var3, "")
                        arrAnalyticalRuns(3, Count1) = NZ(rs.Fields("EXTRACTIONDATE").Value, "No entry")
                        arrAnalyticalRuns(4, Count1) = NZ(rs.Fields("RUNSTARTDATE").Value, "No entry")

                        var2 = "No regression performed"
                        arrAnalyticalRuns(5, Count1) = var2

                        arrAnalyticalRuns(6, Count1) = NZ(rs.Fields("RUNDESCRIPTION").Value, "")
                        arrAnalyticalRuns(7, Count1) = "Not applicable" ' NZ(rs1.Fields("ACCEPTREJECTREASON").Value, "")
                        arrAnalyticalRuns(8, Count1) = NZ(arrAnalytes(1, Count2), "")
                        arrAnalyticalRuns(9, Count1) = rs.Fields("RUNTYPEID").Value
                        arrAnalyticalRuns(10, Count1) = "NA"
                        arrAnalyticalRuns(11, Count1) = "NA"
                        arrAnalyticalRuns(12, Count1) = "NA"  'AssayID
                        arrAnalyticalRuns(13, Count1) = "No"  'boolInthisRunsAssayID
                        arrAnalyticalRuns(14, Count1) = 1 'RUNANALYTEREGRESSIONSTATUS
                        arrAnalyticalRuns(15, Count1) = "NA"

                    Else
                        arrAnalyticalRuns(1, Count1) = int1 'rs.Fields("RUNID").Value
                        'evalutate pagenumber
                        var1 = NZ(rs.Fields("NOTEBOOK").Value, "")
                        var2 = NZ(rs.Fields("PAGENUMBER").Value, "")
                        If InStr(1, var2, var1, vbTextCompare) > 0 Then
                            var3 = var2
                        ElseIf Len(var2) = 0 Then
                            var3 = var1
                        Else
                            var3 = var1 & "-" & var2
                        End If
                        arrAnalyticalRuns(2, Count1) = NZ(var3, "")
                        arrAnalyticalRuns(3, Count1) = NZ(rs.Fields("EXTRACTIONDATE").Value, "No entry")
                        arrAnalyticalRuns(4, Count1) = NZ(rs.Fields("RUNSTARTDATE").Value, "No entry")
                        var1 = NZ(rs1.Fields("RUNANALYTEREGRESSIONSTATUS").Value, 0) '3=Pass, <>3 = Fail
                        str1 = "RUNANALYTEREGRESSIONSTATUS = " & var1
                        arrAnalyticalRuns(14, Count1) = var1 'RUNANALYTEREGRESSIONSTATUS
                        rsRS.Filter = ""
                        rsRS.Filter = str1
                        var2 = rsRS.Fields("ANAREGSTATUSDESC").Value
                        rsRS.Filter = ""
                        arrAnalyticalRuns(5, Count1) = var2
                        arrAnalyticalRuns(6, Count1) = NZ(rs.Fields("RUNDESCRIPTION").Value, "")
                        arrAnalyticalRuns(7, Count1) = NZ(rs1.Fields("ACCEPTREJECTREASON").Value, "")
                        arrAnalyticalRuns(8, Count1) = NZ(arrAnalytes(1, Count2), "")
                        arrAnalyticalRuns(9, Count1) = rs.Fields("RUNTYPEID").Value
                        arrAnalyticalRuns(10, Count1) = NZ(rs1.Fields("NM").Value, "")
                        arrAnalyticalRuns(11, Count1) = NZ(rs1.Fields("VEC").Value, "")

                        '**** Added 20-Jan-2016 by NDL to support an Analytical Run Summary Table that record which sub-Analytes
                        'exist in those runs. *****
                        arrAnalyticalRuns(12, Count1) = NZ(arrAnalytes(2, Count2), "") 'AnalyteID
                        'Search to see if this sub-Analyte is in this Run (i.e. is it in the run's AssayID's AssayAnalytes table)

                        'Find table row which has the analyte and the run
                        Dim strMasterAssayID As String = ""

                        'See if the MasterAssayID of the Analyte in the Run matches the one in the Analyte Assay
                        arrAnalyticalRuns(13, Count1) = "No"
                        'If this analyte is in the same analyte group that is in the run
                        'Filter tblCalStdGroupsAll for this AnalyteID and this RunID
                        Dim dv3 As New DataView(tblCalStdGroupAssayIDsAll)
                        var1 = dv3.Count 'debug
                        strF = "RUNID = " & arrAnalyticalRuns(1, Count1) & " AND ANALYTEID = " & arrAnalyticalRuns(12, Count1)
                        dv3.RowFilter = strF 'This should result in a single row.

                        ''20171219 LEE:
                        ''This logic is incorrect. It is possible to have more than 1 instance of a single AnalyteID for a run if there are multiple matrix
                        'If dv3.Count > 1 Then
                        '    'Debug
                        '    'MsgBox("Issue: There is more than 1 instance of a single AnalyteID for a run")
                        '    var1 = var1
                        'ElseIf dv3.Count = 0 Then 'The run hasn't been calibrated; skip.
                        '    ''Console.WriteLine("Issue: Run " & arrAnalyticalRuns(1, Count1) & " isn't appearing in the tblCalStdGroupAssayIDs for Analyte ID " & arrAnalyticalRuns(12, Count1))
                        '    var1 = var1
                        'Else
                        '    Dim tbldv3 As New DataTable
                        '    tbldv3 = dv3.ToTable("tlbdv3")

                        '    var1 = arrAnalytes(15, Count2)
                        '    var2 = tbldv3.Rows(0).Item("INTGROUP")
                        '    If (arrAnalytes(15, Count2) = tbldv3.Rows(0).Item("INTGROUP")) Then
                        '        arrAnalyticalRuns(13, Count1) = "Yes"
                        '    End If
                        'End If

                        '20171219 LEE: New logic
                        If dv3.Count = 0 Then
                            ''Console.WriteLine("Issue: Run " & arrAnalyticalRuns(1, Count1) & " isn't appearing in the tblCalStdGroupAssayIDs for Analyte ID " & arrAnalyticalRuns(12, Count1))

                            '20171219 LEE:
                            'don't skip
                            'need to evaluate for boolInThisRunsAssayID

                            var1 = var1
                        Else
                            Dim tbldv3 As New DataTable
                            tbldv3 = dv3.ToTable("tlbdv3")

                            'arrAnalyticalRuns
                            '1=RUNID, 2=NOTEBOOK-PAGENUMBER, 3=EXTRACTIONDATE, 4=RUNSTARTDATE, 5=ANAREGSTATUSDESC
                            '6=RUNDESCRIPTION, 7=ACCEPTREJECTREASON, 8=ANALYTE, 9=RUNTYPEID, 10=NM, 11=VEC, 12=ANALYTEID, 13=BOOLINTHISRUNSASSAYID, 14=RUNANALYTEREGRESSIONSTATUS,15=INSTGROUPNAME

                            var1 = arrAnalytes(15, Count2) 'debug
                            var2 = tbldv3.Rows(0).Item("INTGROUP") 'debug
                            var3 = arrAnalyticalRuns(13, Count1)

                            '20180321 LEE: Do not need to re-establish arranalytes(15, anymore
                            'arrAnalytes(15, Count2) = tbldv3.Rows(0).Item("INTGROUP")

                            'If (arrAnalytes(15, Count2) = tbldv3.Rows(0).Item("INTGROUP")) Then
                            If IsDBNull(var2) = False Then
                                arrAnalyticalRuns(13, Count1) = "Yes"
                            End If
                        End If

                        var1 = rs1.Fields("INSTGROUPNAME").Value
                        arrAnalyticalRuns(15, Count1) = NZ(rs1.Fields("INSTGROUPNAME").Value, "NA")

                    End If
                    If rs.EOF Then
                    Else
                        rs.MoveNext()
                    End If
                Next Count3

            Next Count2

            var1 = var1

            ''debug
            'Try
            '    'Console.WriteLine("Start arrAnalyticalRuns")
            '    var1 = "1=RUNID, 2=NOTEBOOK-PAGENUMBER, 3=EXTRACTIONDATE, 4=RUNSTARTDATE, 5=ANAREGSTATUSDESC, 6=RUNDESCRIPTION, 7=ACCEPTREJECTREASON, 8=ANALYTE, 9=RUNTYPEID, 10=NM, 11=VEC"
            '    var2 = Replace(var1, ",", ";", 1, -1, CompareMethod.Text)
            '    'Console.WriteLine(var2)
            '    For Count2 = 1 To ctAnalyticalRuns
            '        var1 = ""
            '        For Count3 = 1 To 13
            '            var2 = arrAnalyticalRuns(Count3, Count2)
            '            var1 = var1 & ";" & var2
            '        Next
            '        'Console.WriteLine(var1)
            '    Next
            '    'Console.WriteLine("End arrAnalyticalRuns")
            'Catch ex As Exception
            '    'Console.WriteLine("End arrAnalyticalRuns")
            'End Try


            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            If rs1.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs1.Close()
            End If

            ReDim arrRunDates(ctAnalyticalRuns * ctAnalytes) ' * ctAnalytes)
            ReDim arrExtDates(ctAnalyticalRuns * ctAnalytes) ' * ctAnalytes)
            'record extraction dates
            For Count1 = 1 To ctAnalyticalRuns * ctAnalytes
                var1 = arrAnalyticalRuns(3, Count1)
                If IsDate(var1) Then
                    arrExtDates(Count1) = arrAnalyticalRuns(3, Count1)
                Else
                    arrExtDates(Count1) = DBNull.Value ' arrAnalyticalRuns(3, Count1)
                End If

            Next
            'record run dates
            For Count1 = 1 To ctAnalyticalRuns * ctAnalytes
                var1 = arrAnalyticalRuns(4, Count1)
                If IsDate(var1) Then
                    arrRunDates(Count1) = arrAnalyticalRuns(4, Count1)
                Else
                    arrRunDates(Count1) = DBNull.Value ' arrAnalyticalRuns(4, Count1)
                End If

            Next

            'determine initial extraction date
            var1 = GetMin(arrRunDates, ctAnalyticalRuns * ctAnalytes) ' * ctAnalytes)
            var2 = GetMin(arrExtDates, ctAnalyticalRuns * ctAnalytes) ' * ctAnalytes)
            'int1 = FindRow("Initial Extraction Date", dgDataWatson)
            int1 = FindRow("Initial Extraction Date", tblWatsonData, "Item")

            str1 = "Retrieving Watson Data...8 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            tblWatsonData.Columns.Item(1).ReadOnly = False
            If IsDate(var2) Then
                var3 = Format(var2, "MMMM dd, yyyy")
            Else
                If IsDate(var1) Then
                    var3 = Format(var1, "MMMM dd, yyyy")
                Else
                    var3 = "NA"
                End If
            End If
            'If IsDate(var1) And IsDate(var2) Then
            '    If var1 < var2 Then
            '        'Sheets("Data").Range("ExtractionInitDate").Offset(0, 1).Value = Format(var1, "mmmm dd, yyyy")
            '        var3 = Format(var1, "MMMM dd, yyyy")
            '        'var3 = Format(var1, LDateFormat)

            '    Else
            '        'Sheets("Data").Range("ExtractionInitDate").Offset(0, 1).Value = Format(var2, "mmmm dd, yyyy")
            '        var3 = Format(var2, "MMMM dd, yyyy")
            '        'var3 = Format(var2, LDateFormat)
            '    End If
            'Else
            '    If IsDate(var2) Then
            '        var3 = Format(var1, "MMMM dd, yyyy")
            '    ElseIf IsDate(var1) Then
            '        var3 = Format(var2, "MMMM dd, yyyy")
            '    Else
            '        var3 = "NA"
            '    End If
            'End If

            tblWatsonData.Rows.Item(int1).Item(1) = var3

            'determine last extraction date
            var1 = GetMax(arrRunDates, ctAnalyticalRuns * ctAnalytes)
            var2 = GetMax(arrExtDates, ctAnalyticalRuns * ctAnalytes)
            int1 = FindRow("Last Extraction Date", tblWatsonData, "Item")
            If IsDate(var2) Then
                var3 = Format(var2, "MMMM dd, yyyy")
            Else
                If IsDate(var1) Then
                    var3 = Format(var1, "MMMM dd, yyyy")
                Else
                    var3 = "NA"
                End If
            End If
            'If IsDate(var1) And IsDate(var2) Then
            '    If var1 > var2 Then
            '        'Sheets("Data").Range("ExtractionCompleteDate").Offset(0, 1).Value = Format(var1, "mmmm dd, yyyy")
            '        var3 = Format(var1, "MMMM dd, yyyy")
            '        'var3 = Format(var1, LDateFormat)
            '    Else
            '        'Sheets("Data").Range("ExtractionCompleteDate").Offset(0, 1).Value = Format(var2, "mmmm dd, yyyy")
            '        var3 = Format(var2, "MMMM dd, yyyy")
            '        'var3 = Format(var2, LDateFormat)
            '    End If
            'Else
            '    If IsDate(var2) Then
            '        var3 = Format(var1, "MMMM dd, yyyy")
            '    ElseIf IsDate(var1) Then
            '        var3 = Format(var2, "MMMM dd, yyyy")
            '    Else
            '        var3 = "NA"
            '    End If
            'End If

            tblWatsonData.Rows.Item(int1).Item(1) = var3


            '******

            'determine first analysis date
            var1 = GetMin(arrRunDates, ctAnalyticalRuns * ctAnalytes) ' * ctAnalytes)
            var2 = GetMin(arrExtDates, ctAnalyticalRuns * ctAnalytes) ' * ctAnalytes)
            'int1 = FindRow("Initial Extraction Date", dgDataWatson)
            int1 = FindRow("Initial Analysis Date", tblWatsonData, "Item")

            str1 = "Retrieving Watson Data...8 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            tblWatsonData.Columns.Item(1).ReadOnly = False
            If IsDate(var1) Then
                var3 = Format(var1, "MMMM dd, yyyy")
            Else
                If IsDate(var2) Then
                    var3 = Format(var2, "MMMM dd, yyyy")
                Else
                    var3 = "NA"
                End If
            End If

            tblWatsonData.Rows.Item(int1).Item(1) = var3

            'determine last analysis date
            var1 = GetMax(arrRunDates, ctAnalyticalRuns * ctAnalytes)
            var2 = GetMax(arrExtDates, ctAnalyticalRuns * ctAnalytes)
            int1 = FindRow("Last Analysis Date", tblWatsonData, "Item")
            If IsDate(var1) Then
                var3 = Format(var1, "MMMM dd, yyyy")
            Else
                If IsDate(var2) Then
                    var3 = Format(var2, "MMMM dd, yyyy")
                Else
                    var3 = "NA"
                End If
            End If

            tblWatsonData.Rows.Item(int1).Item(1) = var3

            '******



            'enter data for Table Analyte IDs
            Dim dvAnalyteGroups As New DataView(tblAnalyteGroups)
            tblAnalyteIDs = dvAnalyteGroups.ToTable("tblAnalyteIDs", True, "ANALYTEID", "ANALYTEDESCRIPTION") 'this should account for matrix and calibrlevels
            tblAnalyteIDs.AcceptChanges()

            'enter data for Table Matrices
            tblMatrices = dvAnalyteGroups.ToTable("tblMatrices", True, "MATRIX")
            tblMatrices.AcceptChanges()

            var1 = tblMatrices.Rows.Count
            var1 = var1 'debug

            var1 = tblAnalyteIDs.Rows.Count
            var1 = var1 'debug

            'enter data in Analytical Run Summary

            'the following for reference
            'arrAnalyticalRuns(2, Count1) = var3
            'arrAnalyticalRuns(3, Count1) = rs.Fields("EXTRACTIONDATE").Value
            'arrAnalyticalRuns(4, Count1) = rs.Fields("RUNSTARTDATE").Value
            'arrAnalyticalRuns(5, Count1) = rs.Fields("RUNANALYTEREGRESSIONSTATUS").Value '3=Pass, 4=Fail
            'arrAnalyticalRuns(6, Count1) = rs.Fields("RUNDESCRIPTION").Value
            'arrAnalyticalRuns(7, Count1) = rs.Fields("ACCEPTREJECTREASON").Value
            'arrAnalyticalRuns(8, Count1) = arrAnalytes(1, Count1)
            'arrAnalyticalRuns(9, Count1) = rs.Fields("RUNTYPEID").Value

            Dim strAnal As String
            Dim drow As DataRow
            Dim ctUniqueAnalyteIDs As Short

            dtbl = tblAnalRunSum
            'remove all rows
            int1 = dtbl.Rows.Count
            int2 = dtbl.Columns.Count
            For Count1 = int1 - 1 To 0 Step -1
                dtbl.Rows.Remove(dtbl.Rows.Item(Count1))
            Next

            'Then, add rows of samples for each run ID, ordered by AnalyteID and RunID (but not
            'StudyDoc Analyte name)
            ctUniqueAnalyteIDs = tblAnalyteIDs.Rows.Count

            '20180713 LEE:
            'This logic will replicate analytes with multiple calibrlevels
            'this is not what we want. We want multiple calibrlevels included, but not replicated
            'so instead redo unique dv

            tblAnalyteIDs = dvAnalyteGroups.ToTable("tblAnalyteIDs", True, "ANALYTEID") 'only do unique analyteid's
            tblAnalyteIDs.AcceptChanges()
            ctUniqueAnalyteIDs = tblAnalyteIDs.Rows.Count

            Dim intAnalyteID As Int64
            Dim strAnalCM As String
            Dim strAnalM As String

            For Count10 = 0 To ctUniqueAnalyteIDs - 1

                intAnalyteID = tblAnalyteIDs.Rows(Count10).Item("AnalyteID")
                strF = "ANALYTEID = " & intAnalyteID
                Dim rowsAARM() As DataRow = tblCalStdGroupAssayIDsAll.Select(strF)
                strAnalM = rowsAARM(0).Item("ANALYTEDESCRIPTION")

                If Count10 = 0 Then 'Don't add blank row before first Analyte
                Else
                    'add a blank row
                    drow = dtbl.NewRow
                    drow.BeginEdit()

                    Try
                        For Count3 = 2 To int2 - 1 'column 0 and 1 are boolean
                            str1 = dtbl.Columns(Count3).ColumnName
                            Select Case str1
                                Case "RUNANALYTEREGRESSIONSTATUS" 'RUNANALYTEREGRESSIONSTATUS is integer that allows null
                                    drow.Item(Count3) = -1
                                Case Else
                                    drow.Item(Count3) = ""
                            End Select

                        Next
                        drow.Item("boolInclude") = False
                        drow.Item("boolIncludeRegr") = False
                        drow.Item("RUNTYPEID") = 1
                        drow.Item("boolInThisRunsAssayID") = "Yes"
                        drow.EndEdit()
                        dtbl.Rows.Add(drow)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                End If

                ''debug
                'var1 = arrAnalytes(16, 1)
                'var1 = var1

                'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET

                'arrAnalyticalRuns
                '1=RUNID, 2=NOTEBOOK-PAGENUMBER, 3=EXTRACTIONDATE, 4=RUNSTARTDATE, 5=ANAREGSTATUSDESC
                '6=RUNDESCRIPTION, 7=ACCEPTREJECTREASON, 8=ANALYTE, 9=RUNTYPEID, 10=NM, 11=VEC, 12=ANALYTEID, 13=BOOLINTHISRUNSASSAYID, 14=RUNANALYTEREGRESSIONSTATUS,15=INSTGROUPNAME

                'Now, go through every analyte and every analytical run, and filter for this analyteID.
                Dim strAnalC As String
                Dim strMatrix As String
                Dim intRunID As Short
                Dim strRunType As String
                Dim boolEx As Boolean

                'debug
                'Dim arrEE(2, 1000)
                'Dim intEE As Int16
                'intEE = 0


                Try

                    For intAR = 1 To ctAnalyticalRuns  'Index for AnlayticalRun

                        boolEx = False

                        intRunID = arrAnalyticalRuns(1, intAR)
                        If intRunID = 24 Then
                            var1 = var1
                        End If

                        '201807813 LEE:

                        strF = "ANALYTEID = " & intAnalyteID & " AND RUNID = " & intRunID
                        Dim rowsAAR() As DataRow = tblCalStdGroupAssayIDsAll.Select(strF)

                        If rowsAAR.Length = 0 Then
                            strAnalC = "NA"
                            strAnal = strAnalM
                            strMatrix = "NA"
                            strRunType = "NA"
                            boolEx = True
                        Else
                            Try
                                strAnalC = rowsAAR(0).Item("ANALYTEDESCRIPTION_C")
                                strAnal = rowsAAR(0).Item("ANALYTEDESCRIPTION")

                                strMatrix = rowsAAR(0).Item("MATRIX")
                                strRunType = rowsAAR(0).Item("RUNTYPE")

                            Catch ex As Exception

                                var1 = ex.Message
                            End Try
                        End If



                        'get matrix
                        'Try
                        '    intRunID = arrAnalyticalRuns(1, intAR)
                        'Catch ex As Exception
                        '    var1 = ex.Message
                        'End Try

                        'strRunType = "UNKNOWNS"
                        'strF = "RUNID = " & intRunID
                        'Dim rowRID() As DataRow = tblAllAnalRuns.Select(strF)
                        'If rowRID.Length = 0 Then
                        '    strMatrix = "Not Assigned"
                        'Else
                        '    strMatrix = rowRID(0).Item("SAMPLETYPEID")
                        '    strRunType = rowRID(0).Item("RUNTYPEDESCRIPTION")
                        'End If

                        drow = dtbl.NewRow()
                        drow.BeginEdit()

                        For Count3 = 0 To dtbl.Columns.Count - 1

                            Select Case Count3
                                Case 0
                                    str1 = "boolInclude"
                                    str2 = ""
                                    boolRO = False
                                Case 1
                                    str1 = "Watson Run ID"
                                    str2 = arrAnalyticalRuns(1, intAR)
                                    boolRO = True
                                Case 2
                                    str1 = "Analyte"
                                    str2 = strAnal
                                    boolRO = True
                                Case 3
                                    str1 = "Notebook ID"
                                    str2 = arrAnalyticalRuns(2, intAR)
                                    boolRO = True
                                Case 4
                                    str1 = "Extraction Date" ' 
                                    var1 = NZ(arrAnalyticalRuns(3, intAR), #1/1/1900#)
                                    If IsDate(var1) Then
                                        var2 = Format(CDate(var1), LDateFormat)
                                    Else
                                        var2 = var1
                                    End If
                                    'var2 = Format(CDate(var1), LDateFormat)
                                    str2 = var2
                                    boolRO = True
                                Case 5
                                    str1 = "Analysis Date"
                                    var1 = NZ(arrAnalyticalRuns(4, intAR), #1/1/1900#)
                                    If IsDate(var1) Then
                                        var2 = Format(CDate(var1), LDateFormat)
                                        'intEE = intEE + 1
                                        'arrEE(1, intEE) = var1
                                        'arrEE(2, intEE) = var2
                                    Else
                                        var2 = var1

                                    End If

                                    str2 = var2
                                    boolRO = True
                                Case 6
                                    str1 = "Pass/Fail"
                                    str2 = arrAnalyticalRuns(5, intAR)
                                    boolRO = True
                                Case 7
                                    str1 = "Samples"
                                    str2 = arrAnalyticalRuns(6, intAR)
                                    boolRO = True
                                Case 8
                                    str1 = "Watson Comments"
                                    str2 = arrAnalyticalRuns(7, intAR)
                                    boolRO = True
                                Case 9
                                    str1 = "User Comments"
                                    str2 = ""
                                    boolRO = False
                                Case 10
                                    str1 = "RUNTYPEID"
                                    str2 = arrAnalyticalRuns(9, intAR)
                                Case 11
                                    str1 = "boolIncludeRegr"
                                    str2 = ""
                                    boolRO = False
                                Case 12
                                    str1 = "LLOQ"
                                    str2 = arrAnalyticalRuns(10, intAR)
                                    boolRO = True
                                Case 13
                                    str1 = "ULOQ"
                                    str2 = arrAnalyticalRuns(11, intAR)
                                    boolRO = True

                                    'NDL 20-Jan-2016 Added AnalyteID and boolInThisRunsAssayID
                                Case 14
                                    str1 = "AnalyteID"
                                    str2 = intAnalyteID 'arrAnalyticalRuns(12, intAR)
                                    boolRO = True
                                Case 15
                                    str1 = "boolInThisRunsAssayID"
                                    str2 = arrAnalyticalRuns(13, intAR)
                                    boolRO = True

                                    '20160205 LEE: Added this column
                                Case 16
                                    str1 = "Analyte_C"
                                    str2 = strAnalC
                                    boolRO = True

                                    '20160205 LEE: Added this column
                                Case 17
                                    str1 = "Matrix"
                                    str2 = strMatrix
                                    boolRO = True

                                    '20160319 LEE: Added this column
                                Case 18
                                    str1 = "Run Type"
                                    str2 = strRunType
                                    boolRO = True

                                    '20160906 LEE: Added this column
                                Case 19
                                    str1 = "RUNANALYTEREGRESSIONSTATUS"
                                    str2 = arrAnalyticalRuns(14, intAR)
                                    boolRO = True

                                Case 20
                                    '20171108 LEE: Added this column for Alturas
                                    str1 = "Instrument ID"
                                    str2 = arrAnalyticalRuns(15, intAR)
                                    boolRO = True

                            End Select

                            If boolEx Then


                                Select Case Count3

                                    Case 10
                                        str2 = "1"
                                    Case 15
                                        str2 = "NO"
                                        'Case 18
                                        '    str2 = "UNKNOWNS"
                                    Case 19
                                        str2 = "0"
                                    Case 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16, 17, 18, 20
                                        str2 = "NA"
                                    Case Else

                                End Select

                            End If

                            Select Case Count3
                                Case 19
                                    Try
                                        int1 = CInt(NZ(str2, 0))
                                        drow.Item(str1) = int1
                                    Catch ex As Exception
                                        var1 = var1
                                    End Try

                                Case Else
                                    If StrComp(str1, "boolInclude", CompareMethod.Text) = 0 Or StrComp(str1, "boolIncludeRegr", CompareMethod.Text) = 0 Then
                                        If boolEx Then
                                            drow.Item(str1) = False
                                        Else
                                            drow.Item(str1) = True
                                        End If

                                    Else
                                        drow.Item(str1) = str2
                                    End If
                            End Select

                        Next Count3

                        drow.EndEdit()
                        dtbl.Rows.Add(drow)


                        'For Count1 = 1 To ctAnalytes 'Index for Analytes

                        '    intARAnalytes = ((Count1 - 1) * ctAnalyticalRuns) + intAR  ' Index for AnalyticalRunAnalytes



                        '    If (arrAnalytes(2, Count1) = intAnalyteID) Then
                        '        'Show this Run if AnalyteID is matches the one in this loop
                        '        'Try
                        '        '    strAnalC = arrAnalytes(1, Count1)
                        '        '    strAnal = arrAnalytes(14, Count1)
                        '        'Catch ex As Exception
                        '        '    var1 = ex.Message
                        '        'End Try

                        '        ''get matrix
                        '        'Try
                        '        '    intRunID = arrAnalyticalRuns(1, intAR)
                        '        'Catch ex As Exception
                        '        '    var1 = ex.Message
                        '        'End Try

                        '        'strRunType = "UNKNOWNS"
                        '        'strF = "RUNID = " & intRunID
                        '        'Dim rowRID() As DataRow = tblAllAnalRuns.Select(strF)
                        '        'If rowRID.Length = 0 Then
                        '        '    strMatrix = "Not Assigned"
                        '        'Else
                        '        '    strMatrix = rowRID(0).Item("SAMPLETYPEID")
                        '        '    strRunType = rowRID(0).Item("RUNTYPEDESCRIPTION")
                        '        'End If

                        '        'drow = dtbl.NewRow()
                        '        'drow.BeginEdit()

                        '        'For Count3 = 0 To dtbl.Columns.Count - 1
                        '        '    Select Case Count3
                        '        '        Case 0
                        '        '            str1 = "boolInclude"
                        '        '            str2 = ""
                        '        '            boolRO = False
                        '        '        Case 1
                        '        '            str1 = "Watson Run ID"
                        '        '            str2 = arrAnalyticalRuns(1, intAR)
                        '        '            boolRO = True
                        '        '        Case 2
                        '        '            str1 = "Analyte"
                        '        '            str2 = strAnal
                        '        '            boolRO = True
                        '        '        Case 3
                        '        '            str1 = "Notebook ID"
                        '        '            str2 = arrAnalyticalRuns(2, intAR)
                        '        '            boolRO = True
                        '        '        Case 4
                        '        '            str1 = "Extraction Date" ' 
                        '        '            var1 = NZ(arrAnalyticalRuns(3, intAR), #1/1/1900#)
                        '        '            If IsDate(var1) Then
                        '        '                var2 = Format(CDate(var1), LDateFormat)
                        '        '            Else
                        '        '                var2 = var1
                        '        '            End If
                        '        '            'var2 = Format(CDate(var1), LDateFormat)
                        '        '            str2 = var2
                        '        '            boolRO = True
                        '        '        Case 5
                        '        '            str1 = "Analysis Date"
                        '        '            var1 = NZ(arrAnalyticalRuns(4, intAR), #1/1/1900#)
                        '        '            If IsDate(var1) Then
                        '        '                var2 = Format(CDate(var1), LDateFormat)
                        '        '                'intEE = intEE + 1
                        '        '                'arrEE(1, intEE) = var1
                        '        '                'arrEE(2, intEE) = var2
                        '        '            Else
                        '        '                var2 = var1

                        '        '            End If

                        '        '            str2 = var2
                        '        '            boolRO = True
                        '        '        Case 6
                        '        '            str1 = "Pass/Fail"
                        '        '            str2 = arrAnalyticalRuns(5, intARAnalytes)
                        '        '            boolRO = True
                        '        '        Case 7
                        '        '            str1 = "Samples"
                        '        '            str2 = arrAnalyticalRuns(6, intAR)
                        '        '            boolRO = True
                        '        '        Case 8
                        '        '            str1 = "Watson Comments"
                        '        '            str2 = arrAnalyticalRuns(7, intARAnalytes)
                        '        '            boolRO = True
                        '        '        Case 9
                        '        '            str1 = "User Comments"
                        '        '            str2 = ""
                        '        '            boolRO = False
                        '        '        Case 10
                        '        '            str1 = "RUNTYPEID"
                        '        '            str2 = arrAnalyticalRuns(9, intAR)
                        '        '        Case 11
                        '        '            str1 = "boolIncludeRegr"
                        '        '            str2 = ""
                        '        '            boolRO = False
                        '        '        Case 12
                        '        '            str1 = "LLOQ"
                        '        '            str2 = arrAnalyticalRuns(10, intARAnalytes)
                        '        '            boolRO = True
                        '        '        Case 13
                        '        '            str1 = "ULOQ"
                        '        '            str2 = arrAnalyticalRuns(11, intARAnalytes)
                        '        '            boolRO = True

                        '        '            'NDL 20-Jan-2016 Added AnalyteID and boolInThisRunsAssayID
                        '        '        Case 14
                        '        '            str1 = "AnalyteID"
                        '        '            str2 = arrAnalyticalRuns(12, intARAnalytes)
                        '        '            boolRO = True
                        '        '        Case 15
                        '        '            str1 = "boolInThisRunsAssayID"
                        '        '            str2 = arrAnalyticalRuns(13, intARAnalytes)
                        '        '            boolRO = True

                        '        '            '20160205 LEE: Added this column
                        '        '        Case 16
                        '        '            str1 = "Analyte_C"
                        '        '            str2 = strAnalC
                        '        '            boolRO = True

                        '        '            '20160205 LEE: Added this column
                        '        '        Case 17
                        '        '            str1 = "Matrix"
                        '        '            str2 = strMatrix
                        '        '            boolRO = True

                        '        '            '20160319 LEE: Added this column
                        '        '        Case 18
                        '        '            str1 = "Run Type"
                        '        '            str2 = strRunType
                        '        '            boolRO = True

                        '        '            '20160906 LEE: Added this column
                        '        '        Case 19
                        '        '            str1 = "RUNANALYTEREGRESSIONSTATUS"
                        '        '            str2 = arrAnalyticalRuns(14, intARAnalytes)
                        '        '            boolRO = True

                        '        '        Case 20
                        '        '            '20171108 LEE: Added this column for Alturas
                        '        '            str1 = "Instrument ID"
                        '        '            str2 = arrAnalyticalRuns(15, intARAnalytes)
                        '        '            boolRO = True

                        '        '    End Select
                        '        '    Select Case Count3
                        '        '        Case 19
                        '        '            Try
                        '        '                int1 = CInt(NZ(str2, 0))
                        '        '                drow.Item(str1) = int1
                        '        '            Catch ex As Exception
                        '        '                var1 = var1
                        '        '            End Try

                        '        '        Case Else
                        '        '            If StrComp(str1, "boolInclude", CompareMethod.Text) = 0 Or StrComp(str1, "boolIncludeRegr", CompareMethod.Text) = 0 Then
                        '        '                drow.Item(str1) = True
                        '        '            Else
                        '        '                drow.Item(str1) = str2
                        '        '            End If
                        '        '    End Select

                        '        'Next
                        '        'drow.EndEdit()
                        '        'dtbl.Rows.Add(drow)

                        '    End If

                        'Next Count1


                        ''debug
                        'If intEE > 0 Then
                        '    var1 = ""
                        '    Dim CountEE As Int16
                        '    For CountEE = 1 To intEE
                        '        var2 = "var1: " & arrEE(1, CountEE) & ", var2: " & arrEE(2, CountEE)
                        '        var1 = var1 & ChrW(10) & var2
                        '    Next
                        '    MsgBox(var1)
                        'End If

                    Next intAR

                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1 'debug
                End Try

            Next Count10


            ''debug
            'var1 = arrAnalytes(16, 1)
            'var1 = var1

            ''debug
            'Console.WriteLine("Start This")
            'For Count2 = 0 To dtbl.Columns.Count - 1
            '    var1 = var1 & ";" & dtbl.Columns(Count2).ColumnName
            'Next
            'Console.WriteLine(var1)
            'For Count3 = 0 To dtbl.Rows.Count - 1
            '    var1 = ""
            '    For Count2 = 0 To dtbl.Columns.Count - 1
            '        var1 = var1 & ";" & dtbl.Rows(Count3).Item(Count2)
            '    Next
            '    Console.WriteLine(var1)
            'Next
            'Console.WriteLine("End This")

            dv = dtbl.DefaultView
            dv.AllowDelete = False
            dv.AllowNew = False
            dv.AllowEdit = True
            'dganalyticalRunSummary.DataSource = dv
            'dganalyticalRunSummary.Refresh()

            frmH.dgvAnalyticalRunSummary.DataSource = dv
            frmH.dgvAnalyticalRunSummary.Refresh()

            'Fill tblAnalyeIDs


            ''''debugWriteLine("10")
            Count1 = 0

            '***End here 4


            str1 = "Retrieving Watson Data...9 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            'frmH.pb1.Value = ctPB
            If ctPB > frmH.pb1.Maximum Then
                frmH.pb1.Maximum = frmH.pb1.Maximum + 100
            End If
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()


            '***Start here 6
            'retrieve Summary of Back Calculated Std Concentration info
            str1 = "Retrieving summary of back calculated calibration standard concentration info..."

            'must gather information per analyte per run
            'must determine number of calibration stds

            ReDim arrRegCon(ctAnalyticalRuns)

            str1 = "Retrieving Watson Data...10" & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'rs.ActiveConnection = Nothing

            'old tblAccAnalRuns ended here. 


            'retrieve calibration std levels info, but must filter for accepted analytical runs
            'complicated studies have replicate assayids
            'must filter for only assay ids
            '
            Dim dvTTc As DataView = New DataView(tblAccAnalRuns)
            Dim tblTTc As DataTable = dvTTc.ToTable("TTc", True, "ASSAYID")
            Dim str5 As String
            'str5 = "("
            'If boolAccess Then
            '    For Count2 = 0 To tblTTc.Rows.Count - 1
            '        var1 = tblTTc.Rows(Count2).Item("ASSAYID")
            '        str5 = str5 & "(ASSAY.ASSAYID) = " & var1 & " OR "
            '    Next
            'Else
            '    For Count2 = 0 To tblTTc.Rows.Count - 1
            '        var1 = tblTTc.Rows(Count2).Item("ASSAYID")
            '        str5 = str5 & "(" & strSchema & ".ASSAY.ASSAYID) = " & var1 & " OR "
            '    Next
            'End If

            ''strip off last OR
            'str5 = Left(str5, Len(str5) - 4)
            'str5 = str5 & ")) "
            '20150911 Larry:  str5 isn't needed anymore, but leave here just in case we need to go back to old code
            '''''''Console.WriteLine("str5: " & str5)

            '''''''''''''''''''''''Console.WriteLine(strSQL)


            'get more assayid stuff: labels for QCs and CalStds

            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYREPS.ASSAYID, ASSAYREPS.LEVELNUMBER, ASSAYREPS.ID, ASSAYREPS.KNOWNTYPE, ASSAYREPS.STUDYID, ASSAYREPS.NUMBEROFREPLICATES "
                str2 = "FROM(ASSAYREPS) "
                str3 = "WHERE (((ASSAYREPS.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ASSAYREPS.ASSAYID, ASSAYREPS.LEVELNUMBER; "
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYREPS.ASSAYID, " & strSchema & ".ASSAYREPS.LEVELNUMBER, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ASSAYREPS.KNOWNTYPE, " & strSchema & ".ASSAYREPS.STUDYID, " & strSchema & ".ASSAYREPS.NUMBEROFREPLICATES "
                str2 = "FROM(" & strSchema & ".ASSAYREPS) "
                str3 = "WHERE (((" & strSchema & ".ASSAYREPS.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYREPS.ASSAYID, " & strSchema & ".ASSAYREPS.LEVELNUMBER; "
            End If

            strSQL = str1 & str2 & str3 & str4

            '''Console.WriteLine(strSQL)

            Dim rsAS As New ADODB.Recordset
            rsAS.CursorLocation = CursorLocationEnum.adUseClient
            rsAS.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsAS.ActiveConnection = Nothing
            If rsAS.EOF And rsAS.BOF Then
            Else
                rsAS.MoveFirst()
            End If

            '''''''''Console.WriteLine("tblAssayLabels: " & strSQL)

            'save this recordset in a datatable
            tblAssayLabels.Clear()
            tblAssayLabels.AcceptChanges()
            tblAssayLabels.BeginLoadData()
            daDoPr.Fill(tblAssayLabels, rsAS)
            tblAssayLabels.EndLoadData()
            rsAS.Close()


            'now do tblBCStdsAssayIDAll
            If boolANSI Then

                If boolAccess Then
                    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                    str2 = "FROM ASSAYANALYTES INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD')) "
                    str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
                Else
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                    str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                    str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD')) "
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"
                End If


            End If

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblBCStdsAssayIDAll : " & strSQL)
            '
            'rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'used to find number of calibration points
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Filter = ""
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            If rs.EOF And rs.BOF Then
            Else
                rs.MoveFirst()
            End If

            'Console.WriteLine("tblBCStdsAssayIDAll: " & strSQL)

            'save this recordset in a datatable
            tblBCStdsAssayIDAll.Clear()
            tblBCStdsAssayIDAll.AcceptChanges()
            tblBCStdsAssayIDAll.BeginLoadData()
            daDoPr.Fill(tblBCStdsAssayIDAll, rs)
            tblBCStdsAssayIDAll.EndLoadData()

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If

            Call FixTableConcentrations(tblBCStdsAssayIDAll)
            'now do tblBCQCStdsAll
            If boolANSI Then

                If boolAccess Then
                    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                    str2 = "FROM ASSAYANALYTES INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC')) "
                    str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
                Else
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                    str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                    str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC')) "
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"
                End If


            End If

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblBCQCStdsAll : " & strSQL)
            '
            'rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'used to find number of calibration points
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.Filter = ""
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            If rs.EOF And rs.BOF Then
            Else
                rs.MoveFirst()
            End If

            '''''''''Console.WriteLine("tblBCQCStdsAll: " & strSQL)

            'save this recordset in a datatable
            tblBCQCStdsAll.Clear()
            tblBCQCStdsAll.AcceptChanges()
            tblBCQCStdsAll.BeginLoadData()
            daDoPr.Fill(tblBCQCStdsAll, rs)
            tblBCQCStdsAll.EndLoadData()
            rs.Close()

            Call FixTableConcentrations(tblBCQCStdsAll)

            'now do tblBCStds
            'WHERE (((ASSAYANALYTEKNOWN.STUDYID)=1717) AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND ((ASSAY.ASSAYID)=28205 Or (ASSAY.ASSAYID)=28408))
            'If boolANSI Then
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID " ', ASSAY.ASSAYID "
            '    str2 = "FROM ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID "
            '    'str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND ((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ")) "
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND " & str5
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'Else
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID " ', ASSAY.ASSAYID "
            '    str2 = "FROM ASSAYANALYTEKNOWN, ASSAY "
            '    str2 = str2 & "WHERE ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID "
            '    str3 = "AND (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND " & str5
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'End If

            'adds analyteid
            If boolANSI Then

                If boolAccess Then
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID "
                    str2 = "FROM ASSAYANALYTES INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND " & str5
                    str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                    '20150911 Larry
                    'get rid of str5
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID "
                    str2 = "FROM ANALYTICALRUN INNER JOIN (ANARUNREGPARAMETERS INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ASSAYANALYTES INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON (ASSAYANALYTEKNOWN.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) ON (ANALYTICALRUNANALYTES.RUNID = ASSAY.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ASSAY.STUDYID)) ON (ANARUNREGPARAMETERS.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) ON (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) "
                    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((ANALYTICALRUN.RUNTYPEID)<>3)) "
                    str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                    '20170928 LEE: Need to add Matrix (SAMPLETYPEID)
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, CONFIGSAMPLETYPES.SAMPLETYPEID "
                    str2 = "FROM (ANALYTICALRUN INNER JOIN (ANARUNREGPARAMETERS INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ASSAYANALYTES INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON (ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTEKNOWN.STUDYID = ASSAY.STUDYID)) ON (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ASSAY.RUNID)) ON (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANARUNREGPARAMETERS.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((ANALYTICALRUN.RUNTYPEID)<>3)) "
                    str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"


                Else
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
                    str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                    str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND " & str5
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                    '20150911 Larry
                    'get rid of str5
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
                    str2 = "FROM " & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANARUNREGPARAMETERS INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ANARUNREGPARAMETERS.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANARUNREGPARAMETERS.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) "
                    str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3)) "
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                    '20170928 LEE: Need to add Matrix (SAMPLETYPEID)
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                    str2 = "FROM (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANARUNREGPARAMETERS INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ASSAY.RUNID)) ON (" & strSchema & ".ANARUNREGPARAMETERS.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ANARUNREGPARAMETERS.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3)) "
                    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                End If

            End If

            'new tblBCStds that has analyte description
            'is goofy with some data
            'If boolANSI Then
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION "
            '    str2 = "FROM (ANARUNANALYTERESULTS INNER JOIN (ASSAYANALYTES INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) ON ANARUNANALYTERESULTS.RUNID = ASSAY.RUNID) INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID "
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD')) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'Else
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION "
            '    str2 = "FROM ANARUNANALYTERESULTS, ASSAYANALYTES, ASSAYANALYTEKNOWN, ASSAY, GLOBALANALYTES "
            '    str2 = str2 & "WHERE (((ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) AND ANARUNANALYTERESULTS.RUNID = ASSAY.RUNID) AND ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID "
            '    str3 = "AND (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='STANDARD')) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

            'End If

            strSQL = str1 & str2 & str3 & str4

            'Console.WriteLine("tblBCStds: " & strSQL)

            'Dim rsBCStds As New ADODB.Recordset
            '
            'rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'used to find number of calibration points
            If rsBCStds.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsBCStds.Close()
            End If
            rsBCStds.CursorLocation = CursorLocationEnum.adUseClient
            rsBCStds.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsBCStds.ActiveConnection = Nothing

            If rsBCStds.EOF And rsBCStds.BOF Then
            Else
                rsBCStds.MoveLast()
                rsBCStds.MoveFirst()
            End If
            var1 = rsBCStds.RecordCount
            'save this recordset in a datatable
            tblBCStds.Clear()
            tblBCStds.AcceptChanges()
            tblBCStds.BeginLoadData()
            daDoPr.Fill(tblBCStds, rsBCStds)
            tblBCStds.EndLoadData()

            var1 = tblBCStds.Rows.Count 'debugging
            Call FixTableConcentrations(tblBCStds)

            'add columns to tblbcstds
            If tblBCStds.Columns.Contains("AnalyteDescription") Then
            Else
                Dim col10 As New DataColumn
                col10.ColumnName = "AnalyteDescription"
                tblBCStds.Columns.Add(col10)
                Dim col10a As New DataColumn
                col10a.ColumnName = "ASSAYID"
                tblBCStds.Columns.Add(col10a)
            End If

            var1 = tblBCStds.Rows.Count




            '*****


            If boolAccess Then
                str1 = "SELECT DISTINCT ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND "
                str2 = "FROM ANALYTICALRUNSAMPLE INNER JOIN (ANALYTICALRUN INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) "
                str3 = "WHERE (((ANALYTICALRUN.RUNTYPEID)>0) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='QC') AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND "
                str2 = "FROM " & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='QC') AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"
            End If

            strSQL = str1 & str2 & str3 & str4

            '" & strAnaRunPeak & ".

            If boolAccess Then
                str1 = "SELECT DISTINCT ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND "
                str2 = "FROM ANALYTICALRUNSAMPLE INNER JOIN (ANALYTICALRUN INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) "
                str3 = "WHERE (((ANALYTICALRUN.RUNTYPEID)>0) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND "
                str2 = "FROM " & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

            End If

            strSQL = str1 & str2 & str3 & str4

            If boolANSI Then
                If boolAccess Then
                    str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ANARUNANALYTERESULTS.ANALYTEINDEX, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                    str2 = "FROM ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) "
                    'str3 = "WHERE(((ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) And ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) = 'STANDARD') And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                    str3 = "WHERE(((ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) > 0) And ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) And ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) = 'STANDARD') And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) > 0)) "
                    'str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
                Else
                    str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, A" & strSchema & ".NARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                    str2 = "FROM " & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) "
                    'str3 = "WHERE(((ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) And ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) = 'STANDARD') And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                    str3 = "WHERE(((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0) And ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) And ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) = 'STANDARD') And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) > 0)) "
                    'str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
                End If


            Else

            End If

            If boolANSI Then

                If boolAccess Then
                    'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER "
                    str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                    str3 = "WHERE(((ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) > 0) And ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) And ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) = 'STANDARD') And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) > 0)) "
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                    '20160207 LEE: optimized query
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER "
                    str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                    str3 = "WHERE (((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                    '20160223 LEE: Added DECISIONREASON
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANARUNPEAKDECISION.DECISIONREASON "
                    str2 = "FROM ANARUNPEAKDECISION RIGHT JOIN (ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID)) ON (ANARUNPEAKDECISION.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANARUNPEAKDECISION.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANARUNPEAKDECISION.RUNID = " & strAnaRunPeak & ".RUNID) "
                    str3 = "WHERE (((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                    '20170728 LEE: Added SAMPLETYPEID (matrix)
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANARUNPEAKDECISION.DECISIONREASON, CONFIGSAMPLETYPES.SAMPLETYPEID "
                    str2 = "FROM (ANARUNPEAKDECISION RIGHT JOIN (ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (ANARUNPEAKDECISION.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANARUNPEAKDECISION.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANARUNPEAKDECISION.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"


                    '20171124 LEE:
                    'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANARUNPEAKDECISION.DECISIONREASON, CONFIGSAMPLETYPES.SAMPLETYPEID "
                    str2 = "FROM (ANARUNPEAKDECISION RIGHT JOIN (ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (ANARUNPEAKDECISION.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANARUNPEAKDECISION.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANARUNPEAKDECISION.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                Else
                    'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER "
                    str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                    str3 = "WHERE(((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0) And ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) And ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) = 'STANDARD') And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) > 0)) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                    '20160207 LEE: optimized query
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER "
                    str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.STUDYID =" & strSchema & ". ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                    str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                    '20160223 LEE: Added DECISIONREASON
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANARUNPEAKDECISION.DECISIONREASON "
                    str2 = "FROM " & strSchema & ".ANARUNPEAKDECISION RIGHT JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANARUNPEAKDECISION.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNPEAKDECISION.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) "
                    str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY " & strSchema & ". ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                    '20170728 LEE: Added SAMPLETYPEID (matrix)
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANARUNPEAKDECISION.DECISIONREASON, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                    str2 = "FROM (" & strSchema & ".ANARUNPEAKDECISION RIGHT JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".ANARUNPEAKDECISION.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANARUNPEAKDECISION.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNPEAKDECISION.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                    '20171124 LEE:
                    'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, ROUND(" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR," & intDFDec & ") AS ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANARUNPEAKDECISION.DECISIONREASON, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                    str2 = "FROM (" & strSchema & ".ANARUNPEAKDECISION RIGHT JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".ANARUNPEAKDECISION.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANARUNPEAKDECISION.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNPEAKDECISION.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND)='STANDARD') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0)) "
                    str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"



                End If


            End If

            'peak

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblBCStdConcs: " & strSQL)
            ''
            'rs2.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'used to find Back Calculated Standard Concentrations for all accepted analytical runs
            If rs2.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs2.Close()
            End If
            rs2.CursorLocation = CursorLocationEnum.adUseClient

            Try
                rs2.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                var1 = var1
            End Try


            rs2.ActiveConnection = Nothing


            tblBCStdConcs.Clear()
            tblBCStdConcs.AcceptChanges()
            tblBCStdConcs.BeginLoadData()
            daDoPr.Fill(tblBCStdConcs, rs2)
            tblBCStdConcs.EndLoadData()


            'rs2.MoveFirst()

            str1 = "Retrieving Watson Data...11 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'rs2.ActiveConnection = Nothing

            int10 = FindRow("Calibration Levels", tblWatsonAnalRefTable, "Item")
            int20 = FindRow("Minimum r^2", tblWatsonAnalRefTable, "Item")
            int30 = FindRow("Analyte Mean Accuracy Min", tblWatsonAnalRefTable, "Item")
            int40 = FindRow("Analyte Mean Accuracy Max", tblWatsonAnalRefTable, "Item")
            int50 = FindRow("Analyte Precision Min", tblWatsonAnalRefTable, "Item")
            int60 = FindRow("Analyte Precision Max", tblWatsonAnalRefTable, "Item")

            Dim ctCalibrStds As Short
            tblQCStds.Clear()
            tblQCStds.AcceptChanges()

            'tblQCStds used in function ReturnQCStds, that's it
            'ReturnQCStds is used in SearchReplace for [QCSTANDARDLIST]

            For Count1 = 1 To ctAnalytes

                Dim arrBCStds(2, 50) '1=LevelNumber, 2=Concentration
                Dim arrBCStdConcs(4, 50) '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag
                Dim arr1()
                Dim arrAcc() As Double
                Dim arrPrec() As Double

                'Count2 = 0

                ''find number of calibration points
                'rsBCStds.Filter = ""
                ''this is not a good filter
                ''masterassayid is bad
                ''str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1)
                ''1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                ''4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                ''10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET

                ''str1 = "ANALYTEID = " & arrAnalytes(2, Count1)
                ''20170928 LEE: Need to filter for matrix in case of multiple matrix studies
                'var1 = arrAnalytes(16, Count1) 'debug
                'str2 = arrAnalytes(16, Count1)
                'If Len(str2) = 0 Then
                '    str1 = "ANALYTEID = " & arrAnalytes(2, Count1)
                'Else
                '    str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                'End If

                ' '''''debugwriteline("str1: ")
                ' '''''debugwriteline(str1)
                ''
                'rsBCStds.Filter = str1
                'var1 = rsBCStds.RecordCount 'debug
                'var1 = var1 'debug
                'Do Until rsBCStds.EOF
                '    Count2 = Count2 + 1
                '    arrBCStds(1, Count2) = rsBCStds.Fields("LevelNumber").Value
                '    var1 = rsBCStds.Fields("CONCENTRATION").Value
                '    arrBCStds(2, Count2) = rsBCStds.Fields("CONCENTRATION").Value
                '    rsBCStds.MoveNext()
                'Loop
                'ctCalibrStds = Count2

                '20181129 LEE:
                'get number of calibrstds from tblCalStdGroupsAll
                var1 = CInt(NZ(arrAnalytes(15, Count1), -1)) 'intgroup
                Dim rows1() As DataRow = tblCalStdGroupsAcc.Select("INTGROUP = " & var1)
                For Count2 = 0 To rows1.Length - 1
                    arrBCStds(1, Count2) = rows1(Count2).Item("LevelNumber")
                    var1 = CDec(NZ(rows1(Count2).Item("CONCENTRATION"), 0))
                    arrBCStds(2, Count2) = var1
                Next
                ctCalibrStds = rows1.Length

                'record # of Calibration Stds
                tblWatsonAnalRefTable.Rows.Item(int10).Item(Count1) = ctCalibrStds
                'Sheets("AnalRefTables").Range("CalibrationPointNumber").Offset(0, Count1).Value = ctCalibrStds

                ''do this after getting tblRegCon
                ''get min regression r2
                ''int2 = arrAnalytes(3, Count1) 'analyteindex
                'Count2 = 0
                ''1=RUNID, 2=AnalyteIndex, 3=REGRESSIONPARAMETERID(1=Slope, 2=YInt, 3=R2),4=PARAMETERVALUE
                ''1=RUNID,  2=Slope, 3=YInt, 4=R2
                'rsAAR.Filter = ""
                'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1)
                'rsAAR.Filter = str1
                'Do Until rsAAR.EOF
                '    Count2 = Count2 + 1
                '    If Count2 > UBound(arrRegCon) Then
                '        ReDim Preserve arrRegCon(UBound(arrRegCon) + 50)
                '    End If
                '    arrRegCon(Count2) = rsAAR.Fields("RSQUARED").Value
                '    rsAAR.MoveNext()
                'Loop
                ''record R_2
                ''str1 = "0."
                ''For Count2 = 1 To LR2SigFigs
                ''    str1 = str1 & "0"
                ''Next

                'var3 = GetMin(arrRegCon, Count2)
                'var2 = SigFigOrDecString(var3, LR2SigFigs, False)
                'str1 = GetRegrDecStr(LR2SigFigs)
                'var1 = Format(CDec(var2), str1)
                'tblWatsonAnalRefTable.Rows.Item(int20).Item(Count1) = var1

                ReDim arrBCStdConcs(4, (ctAnalyticalRuns * 2 * ctCalibrStds) + 10) '1=LevelNumber, 2=Concentration, 3=RunID
                '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag
                'find Back Calculated Standard Concentrations for all accepted analytical runs
                Count2 = 0
                rs2.Filter = ""
                rs2.Sort = ""
                str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3"
                'must account for matrix
                str2 = arrAnalytes(16, Count1)
                'If Len(str2) = 0 Then
                '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3"
                'Else
                '    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & str2 & "'"
                'End If
                '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
                '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET
                If Len(str2) = 0 Then
                    str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3"
                Else
                    str1 = "ANALYTEID = " & arrAnalytes(2, Count1) & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND SAMPLETYPEID = '" & str2 & "'"
                End If
                'str2 = "RUNID ASC, ASSAYLEVEL ASC"
                'str2 = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
                str2 = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                rs2.Filter = str1
                rs2.Sort = str2

                Dim intctRS2 As Short
                intctRS2 = rs2.RecordCount
                ReDim arrBCStdConcs(4, intctRS2) '1=LevelNumber, 2=Concentration, 3=RunID

                Count2 = 0
                Do Until rs2.EOF
                    If ctCalibrStds = 0 Then
                        Exit Do
                    End If
                    For Count3 = 1 To ctCalibrStds
                        If rs2.EOF Then
                            Exit For
                        End If
                        Count2 = Count2 + 1
                        arrBCStdConcs(1, Count2) = Count3 'LevelNumber
                        Count2 = Count2
                        var1 = rs2.Fields("CONCENTRATION").Value
                        If IsDBNull(rs2.Fields("CONCENTRATION").Value) Then
                            'num1 = rs2.Fields("CONCENTRATION").Value
                            arrBCStdConcs(2, Count2) = rs2.Fields("CONCENTRATION").Value 'Concentration
                        Else
                            num1 = rs2.Fields("CONCENTRATION").Value
                            num2 = NZ(rs2.Fields("ALIQUOTFACTOR").Value, 1)
                            num3 = CDbl(num1 / num2)
                            num1 = SigFigOrDec(num3, LSigFig, False)
                            arrBCStdConcs(2, Count2) = num1 'Concentration
                        End If

                        arrBCStdConcs(3, Count2) = rs2.Fields("RUNID").Value 'RunID
                        arrBCStdConcs(4, Count2) = rs2.Fields("ELIMINATEDFLAG").Value 'Eliminated Flag
                        rs2.MoveNext()
                    Next
                Loop

                inttemprows = Count2

                ReDim arrBCStdActual(inttemprows)
                ReDim arrPrec(ctCalibrStds)
                ReDim arrAcc(ctCalibrStds)

                Try

                    For Count3 = 1 To ctCalibrStds
                        int1 = 0
                        For Count5 = 1 To inttemprows
                            var2 = arrBCStdConcs(1, Count5) 'level
                            var3 = arrBCStdConcs(3, Count5) 'runid
                            'If CInt(var2) = Count3 And CInt(var3) = Count2 Then
                            If CInt(var2) = Count3 Then
                                var1 = arrBCStdConcs(4, Count5)
                                If StrComp(var1, "Y", vbTextCompare) = 0 Or IsDBNull(arrBCStdConcs(2, Count5)) Then 'exclude value
                                Else
                                    int1 = int1 + 1
                                    var7 = arrBCStdConcs(2, Count5)
                                    arrBCStdActual(int1) = arrBCStdConcs(2, Count5)

                                End If
                            Else
                            End If
                        Next

                        'determine stats
                        '20150815 Larry: Don't do this anymore
                        '20160107 Larry: Put back. Messes with existing code to retrieve mins/maxes

                        numMean = SigFigOrDec(Mean(int1, arrBCStdActual), LSigFig, False)
                        numSD = SigFigOrDec(StdDev(int1, arrBCStdActual), LSigFig, False)
                        If numMean = 0 Then
                            arrPrec(Count3) = 0 ' CDec(Format(RoundToDecimal(numSD / numMean * 100, 10), "0.0"))
                        Else
                            arrPrec(Count3) = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(numSD / numMean * 100, intQCDec + 1), intQCDec), strQCDec))
                        End If
                        var3 = arrPrec(Count3)
                        var1 = arrBCStds(2, Count3)
                        If IsDBNull(arrBCStds(2, Count3)) Or arrBCStds(2, Count3) = 0 Then
                            arrAcc(Count3) = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(((numMean / 1) - 1) * 100, intQCDec + 1), intQCDec), strQCDec))
                        Else
                            var4 = arrBCStds(2, Count3) 'debugging
                            arrAcc(Count3) = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(((numMean / arrBCStds(2, Count3)) - 1) * 100, intQCDec + 1), intQCDec), strQCDec))
                        End If
                        ''''''debugwriteline(StdDev(int1, arrBCStdActual) & ";")
                        var2 = arrAcc(Count3)
                        var1 = var1
                    Next Count3
                Catch ex As Exception
                    var4 = ex.Message
                End Try

                var1 = Format(CDec(GetMin(arrAcc, ctCalibrStds)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                var1 = Format(CDec(GetMax(arrAcc, ctCalibrStds)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                var1 = Format(CDec(GetMin(arrPrec, ctCalibrStds)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                var1 = Format(CDec(GetMax(arrPrec, ctCalibrStds)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1

                'legend:
                'int30 = FindRow("Analyte Mean Accuracy Min", tblWatsonAnalRefTable, "Item")
                'int40 = FindRow("Analyte Mean Accuracy Max", tblWatsonAnalRefTable, "Item")
                'int50 = FindRow("Analyte Precision Min", tblWatsonAnalRefTable, "Item")
                'int60 = FindRow("Analyte Precision Max", tblWatsonAnalRefTable, "Item")


            Next Count1
            'dgWatsonAnalRef.Refresh()

            If rsBCStds.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsBCStds.Close()
            End If

            If rs1.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs1.Close()
            End If

            If rs2.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs2.Close()
            End If

            '***End here 6

            str1 = "Retrieving Watson Data...12 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            If ctPB > frmH.pb1.Maximum Then
                ctPB = 1
            End If
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()


            '***Start here 7
            'retrieve Summary of Interpolated QC Concentration info
            str1 = "Retrieving summary of interpolated QC concentration info..."

            'must gather information per analyte per run
            Dim rsF As New ADODB.Recordset
            Dim rsF1 As New ADODB.Recordset
            Dim ctQCAI As Short

            '', ASSAY.ASSAYID
            ''must determine number number of QC levels
            'If boolANSI Then
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT "
            '    'str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT, ASSAY.ASSAYID "
            '    str2 = "FROM ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID)<> 3 Or (ANALYTICALRUN.RUNTYPEID)=2) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'Else
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT "
            '    'str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT, ASSAY.ASSAYID "
            '    str2 = "FROM ANALYTICALRUNANALYTES, ANALYTICALRUN, ASSAYREPS, ASSAYANALYTEKNOWN, ASSAY "
            '    str2 = str2 & "WHERE (((ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            '    str3 = "AND (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID)<> 3 Or (ANALYTICALRUN.RUNTYPEID)=2) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'End If

            'If boolANSI Then

            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT, ASSAYANALYTES.ANALYTEID "
            '    str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) ON (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) "
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID)<> 3 Or (ANALYTICALRUN.RUNTYPEID)=2) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

            'End If

            'removed runtypeid. Get replicates in levels if runtypeid is shown
            If boolAccess Then

                str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAYANALYTES.ANALYTEID "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) ON (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID)<> 3 Or (ANALYTICALRUN.RUNTYPEID)=2) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20170928 LEE: Need to SAMPLETYPEID (Matrix) for multiple maxtrix studies
                str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAYANALYTES.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((ANALYTICALRUN.RUNTYPEID)<>3 Or (ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"


            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<> 3 Or (" & strSchema & ".ANALYTICALRUN.RUNTYPEID)=2) AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20170928 LEE: Need to SAMPLETYPEID (Matrix) for multiple maxtrix studies
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3 Or (" & strSchema & ".ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

            End If
            'sampletypeid
            'legend
            'RUNTYPEDESCRIPTION(RUNTYPEID)
            'UNKNOWNS(1)
            'VALIDATION(2)
            'PSAE(3)
            'MANDATORY REPEATS	4
            'RECOVERY(5)

            'don't exclude PSAE
            'NO! EXCLUDE PSAE
            'If boolANSI Then
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT "
            '    'str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT, ASSAY.ASSAYID "
            '    str2 = "FROM ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID) > 0 Or (ANALYTICALRUN.RUNTYPEID)=2) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'Else
            '    str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT "
            '    'str1 = "SELECT DISTINCT ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS,  ASSAYREPS.FLAGPERCENT, ASSAY.ASSAYID "
            '    str2 = "FROM ANALYTICALRUNANALYTES, ANALYTICALRUN, ASSAYREPS, ASSAYANALYTEKNOWN, ASSAY "
            '    str2 = str2 & "WHERE (((ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) "
            '    str3 = "AND (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID) > 0 Or (ANALYTICALRUN.RUNTYPEID)=2) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'End If

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblBCQCs: " & strSQL)

            'Dim rs4 As New ADODB.Recordset
            'rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            If rs4.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs4.Close()
            End If
            rs4.CursorLocation = CursorLocationEnum.adUseClient
            rs4.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '''''''''''''''''''''''''''''Console.WriteLine("tblBCQCs/rs4: " & strSQL)
            rs4.ActiveConnection = Nothing

            tblBCQCs.Clear()
            tblBCQCs.AcceptChanges()
            tblBCQCs.BeginLoadData()
            daDoPr.Fill(tblBCQCs, rs4)
            tblBCQCs.EndLoadData()
            If rs4.EOF And rs4.BOF Then
            Else
                rs4.MoveFirst()
            End If

            Call FixTableConcentrations(tblBCQCs)

            '****
            'make tblBCQCsAssayID for Groups, MASTERASSAYID is no longer included
            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT, ANALYTICALRUN.RUNID "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((ANALYTICALRUN.RUNTYPEID)<>3 Or (ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160318 LEE: Added , ASSAYREPS.DILUTIONFACTOR, changed to ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ANALYTICALRUN.RUNID, ASSAYREPS.DILUTIONFACTOR "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((ANALYTICALRUN.RUNTYPEID)<>3 Or (ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20170126 LEE: Added , ASSAYREPS.FLAGPERCENT
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ANALYTICALRUN.RUNID, ASSAYREPS.DILUTIONFACTOR, ASSAYREPS.FLAGPERCENT "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((ANALYTICALRUN.RUNTYPEID)<>3 Or (ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                'ACTUALLY Round([ASSAYREPS].[DILUTIONFACTOR]," & intDFDec & ") AS DILUTIONFACTOR
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ANALYTICALRUN.RUNID, Round([ASSAYREPS].[DILUTIONFACTOR]," & intDFDec & ") AS DILUTIONFACTOR, ASSAYREPS.FLAGPERCENT "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((ANALYTICALRUN.RUNTYPEID)<>3 Or (ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.ASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYREPS.FLAGPERCENT, " & strSchema & ".ANALYTICALRUN.RUNID "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3 Or (" & strSchema & ".ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160318 LEE: Added , ASSAYREPS.DILUTIONFACTOR
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYREPS.DILUTIONFACTOR "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3 Or (" & strSchema & ".ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20170126 LEE: Added , ASSAYREPS.FLAGPERCENT
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYREPS.DILUTIONFACTOR, " & strSchema & ".ASSAYREPS.FLAGPERCENT "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3 Or (" & strSchema & ".ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"


                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                'ACTUALLY Round([ASSAYREPS].[DILUTIONFACTOR]," & intDFDec & ") AS DILUTIONFACTOR
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ANALYTICALRUN.RUNID, ROUND(" & strSchema & ".ASSAYREPS.DILUTIONFACTOR," & intDFDec & ") AS DILUTIONFACTOR, " & strSchema & ".ASSAYREPS.FLAGPERCENT "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<>3 Or (" & strSchema & ".ANALYTICALRUN.RUNTYPEID)=2)) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"


            End If

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblBCQCsAssayID: " & strSQL)

            Dim rs11 As New ADODB.Recordset
            If rs11.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs11.Close()
            End If

            str1 = "Before rs1"
            rs11.CursorLocation = CursorLocationEnum.adUseClient
            Try
                rs11.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

            '''''''''''''''''''''''''''''Console.WriteLine("tblBCQCs/rs4: " & strSQL)
            rs11.ActiveConnection = Nothing

            str1 = "Before tblBCQCsAssayID"
            tblBCQCsAssayID.Clear()
            tblBCQCsAssayID.AcceptChanges()
            tblBCQCsAssayID.BeginLoadData()
            daDoPr.Fill(tblBCQCsAssayID, rs11)
            tblBCQCsAssayID.EndLoadData()

            If rs11.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs11.Close()
            End If

            rs11 = Nothing

            Call FixTableConcentrations(tblBCQCsAssayID)

            '****

            ', ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAY.ASSAYID
            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ANALYTICALRUN.RUNID, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAY.ASSAYID "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC')) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ANALYTICALRUN.RUNID, ASSAYANALYTEKNOWN.LEVELNUMBER; "
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYREPS.FLAGPERCENT, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAY.ASSAYID "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC')) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER; "

            End If

            strSQL = str1 & str2 & str3 & str4
            ''Console.WriteLine("tblQCRunIDs: " & strSQL)

            Dim rs41 As New ADODB.Recordset
            'rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            If rs41.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs41.Close()
            End If
            rs41.CursorLocation = CursorLocationEnum.adUseClient
            rs41.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '''Console.WriteLine("tblQCRunIDs: " & strSQL)
            rs41.ActiveConnection = Nothing

            tblQCRunIDs.Clear()
            tblQCRunIDs.AcceptChanges()
            tblQCRunIDs.BeginLoadData()
            daDoPr.Fill(tblQCRunIDs, rs41)
            tblQCRunIDs.EndLoadData()

            If rs41.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs41.Close()
            End If
            rs41 = Nothing

            Call FixTableConcentrations(tblQCRunIDs)

            str1 = "Retrieving Watson Data...13 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()


            'appropriate assayid's (AI)
            If boolANSI Then
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAY.MASTERASSAYID  "
                str2 = "FROM ASSAY INNER JOIN ((ANARUNREGPARAMETERS INNER JOIN ANALYTICALRUN ON (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) ON ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((ANALYTICALRUN.RUNTYPEID)<> 3) AND ((ANALYTICALRUN.RUNSTATUS)=3 Or (ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID;"
            Else
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAY.MASTERASSAYID  "
                str2 = "FROM ASSAY, ANARUNREGPARAMETERS, ANALYTICALRUN, ASSAYANALYTEKNOWN "
                str2 = str2 & "WHERE (((ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) AND ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "AND (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((ANALYTICALRUN.RUNTYPEID)<> 3) AND ((ANALYTICALRUN.RUNSTATUS)=3 Or (ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID;"
            End If

            'Include PSAE
            If boolANSI Then
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ASSAY INNER JOIN ((ANARUNREGPARAMETERS INNER JOIN ANALYTICALRUN ON (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) ON ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUN.RUNSTATUS)=3 Or (ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID;"
            Else
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ASSAY, ANARUNREGPARAMETERS, ANALYTICALRUN, ASSAYANALYTEKNOWN "
                str2 = str2 & "WHERE (((ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) AND ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "AND (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUN.RUNSTATUS)=3 Or (ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID;"
            End If

            'Add ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT
            'Include PSAE
            If boolANSI Then
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                str2 = "FROM ASSAY INNER JOIN ((ANARUNREGPARAMETERS INNER JOIN ANALYTICALRUN ON (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) ON ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUN.RUNSTATUS)=3 Or (ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID;"
            Else
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                str2 = "FROM ASSAY, ANARUNREGPARAMETERS, ANALYTICALRUN, ASSAYANALYTEKNOWN "
                str2 = str2 & "WHERE (((ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) AND ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "AND (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUN.RUNSTATUS)=3 Or (ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID;"
            End If

            If boolAccess Then
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                str2 = "FROM ASSAY INNER JOIN ((ANARUNREGPARAMETERS INNER JOIN ANALYTICALRUN ON (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) ON ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUN.RUNSTATUS)=3 Or (ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT "
                str2 = "FROM " & strSchema & ".ASSAY INNER JOIN ((" & strSchema & ".ANARUNREGPARAMETERS INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANARUNREGPARAMETERS.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANARUNREGPARAMETERS.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON " & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) ON " & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID)=1) AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0) AND ((" & strSchema & ".ANALYTICALRUN.RUNSTATUS)=3 Or (" & strSchema & ".ANALYTICALRUN.RUNSTATUS)=7)) "
                str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID;"
            End If

            strSQL = str1 & str2 & str3 & str4

            'Console.WriteLine("tblQCAI:  " & strSQL)

            'legend
            'RUNTYPEDESCRIPTION(RUNTYPEID)
            'UNKNOWNS(1)
            'VALIDATION(2)
            'PSAE(3)
            'MANDATORY REPEATS	4
            'RECOVERY(5)

            If rs1.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs1.Close()
            End If
            rs1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs1.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '''''''''''''''''''''''''''''''''''''Console.WriteLine(strSQL)
            rs1.ActiveConnection = Nothing

            'add columns to table tblQCAI
            tblQCAI.Clear()
            tblQCAI.AcceptChanges()
            tblQCAI.BeginLoadData()
            daDoPr.Fill(tblQCAI, rs1)
            tblQCAI.EndLoadData()

            FixTableConcentrations(tblQCAI)

            If rs1.EOF And rs1.BOF Then
            Else
                rs1.MoveFirst()
            End If
            'rs1.ActiveConnection = Nothing

            str1 = "Retrieving Watson Data...14 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            '#ofReplicates for each level of QC
            If boolANSI Then
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.ANALYTEINDEX " ', ANALYTICALRUNSAMPLE.ALIQUOTFACTOR "
                str2 = "FROM ASSAYANALYTEKNOWN INNER JOIN ((ASSAY INNER JOIN ANALYTICALRUN ON ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUNANALYTES.RUNID)) ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<> 3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.REPLICATENUMBER;"
            Else
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.ANALYTEINDEX " ', ANALYTICALRUNSAMPLE.ALIQUOTFACTOR "
                'str2 = "FROM ASSAYANALYTEKNOWN, ANALYTICALRUNSAMPLE, ASSAY, ANALYTICALRUN, ANALYTICALRUNANALYTES "
                str2 = "FROM ASSAYANALYTEKNOWN, ANALYTICALRUN, ASSAY ,ANALYTICALRUNSAMPLE, ANALYTICALRUNANALYTES "
                str2 = str2 & "WHERE ((ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND ((ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUNANALYTES.RUNID)) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) AND ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID "
                str3 = "AND (((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<> 3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.REPLICATENUMBER;"
            End If

            'include PSAE
            If boolANSI Then
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.ANALYTEINDEX, ANALYTICALRUN.RUNTYPEID " ', ANALYTICALRUNSAMPLE.ALIQUOTFACTOR "
                str2 = "FROM ASSAYANALYTEKNOWN INNER JOIN ((ASSAY INNER JOIN ANALYTICALRUN ON ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUNANALYTES.RUNID)) ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.REPLICATENUMBER;"
            Else
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.ANALYTEINDEX, ANALYTICALRUN.RUNTYPEID " ', ANALYTICALRUNSAMPLE.ALIQUOTFACTOR "
                'str2 = "FROM ASSAYANALYTEKNOWN, ANALYTICALRUNSAMPLE, ASSAY, ANALYTICALRUN, ANALYTICALRUNANALYTES "
                str2 = "FROM ASSAYANALYTEKNOWN, ANALYTICALRUN, ASSAY ,ANALYTICALRUNSAMPLE, ANALYTICALRUNANALYTES "
                str2 = str2 & "WHERE ((ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND ((ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUNANALYTES.RUNID)) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID)) AND ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID "
                str3 = "AND (((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.REPLICATENUMBER;"
            End If

            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ASSAY.MASTERASSAYID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUNSAMPLE.STUDYID, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.ANALYTEINDEX, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ((ASSAY INNER JOIN ANALYTICALRUN ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID)) INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) ON (ASSAYANALYTEKNOWN.STUDYID = ASSAY.STUDYID) AND (ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.REPLICATENUMBER;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUNSAMPLE.REPLICATENUMBER, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN ((" & strSchema & ".ASSAY INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID)) INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) ON (" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.REPLICATENUMBER;"
            End If

            strSQL = str1 & str2 & str3 & str4
            '''Console.WriteLine("tblQCReps: " & strSQL)
            ''
            'rs2.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'Dim rs20 As New ADODB.Recordset
            'must be client side because will filter later
            rs20.CursorLocation = CursorLocationEnum.adUseClient
            rs20.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs20.ActiveConnection = Nothing
            'add columns to table tblQCReps
            tblQCReps.Clear()
            tblQCReps.AcceptChanges()
            tblQCReps.BeginLoadData()
            daDoPr.Fill(tblQCReps, rs20)
            tblQCReps.EndLoadData()

            FixTableConcentrations(tblQCReps)

            If rs20.EOF And rs20.BOF Then
            Else
                rs20.MoveFirst()
            End If
            var1 = rs20.RecordCount
            'rs2.ActiveConnection = Nothing

            str1 = "Retrieving Watson Data...15 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()


            'QC Results
            '*****setup rs for finding nominal concentration
            'Hmmm. This is the same as rs4 above, except also retrieving ANALYTICALRUN.ASSAYID and WHERE is different
            'If boolANSI Then
            '    str1 = "SELECT DISTINCT ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT "
            '    str2 = "FROM ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) "
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID)<> 3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'Else
            '    str1 = "SELECT DISTINCT ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT "
            '    str2 = "FROM ANALYTICALRUNANALYTES, ANALYTICALRUN, ASSAYREPS, ASSAYANALYTEKNOWN, ASSAY "
            '    str2 = str2 & "WHERE (((ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) "
            '    str3 = "AND (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID)<> 3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

            'End If

            'need PSAE
            'If boolANSI Then
            '    str1 = "SELECT DISTINCT ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT "
            '    str2 = "FROM ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) "
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

            'Else
            '    str1 = "SELECT DISTINCT ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT "
            '    str2 = "FROM ANALYTICALRUNANALYTES, ANALYTICALRUN, ASSAYREPS, ASSAYANALYTEKNOWN, ASSAY "
            '    str2 = str2 & "WHERE (((ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) "
            '    str3 = "AND (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
            '    str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"
            'End If

            If boolAccess Then
                str1 = "SELECT DISTINCT ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYREPS.FLAGPERCENT, ASSAYANALYTES.ANALYTEID "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160318 LEE: changed to  ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT
                str1 = "SELECT DISTINCT ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER, ASSAYANALYTEKNOWN.CONCENTRATION, ASSAYANALYTEKNOWN.STUDYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYREPS.ID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, ASSAYANALYTES.ANALYTEID "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ASSAYREPS INNER JOIN (ASSAYANALYTEKNOWN INNER JOIN ASSAY ON ASSAYANALYTEKNOWN.ASSAYID = ASSAY.ASSAYID) ON (ASSAYREPS.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYREPS.KNOWNTYPE = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYREPS.LEVELNUMBER = ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ASSAY.MASTERASSAYID, ASSAYANALYTEKNOWN.ANALYTEINDEX, ASSAYANALYTEKNOWN.LEVELNUMBER;"

            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYREPS.FLAGPERCENT, " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0) AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

                '20160318 LEE: changed to  ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYREPS.ID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT, " & strSchema & ".ASSAYANALYTES.ANALYTEID "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ASSAYREPS INNER JOIN (" & strSchema & ".ASSAYANALYTEKNOWN INNER JOIN " & strSchema & ".ASSAY ON " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) ON (" & strSchema & ".ASSAYREPS.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYREPS.KNOWNTYPE = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYREPS.LEVELNUMBER = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE)='QC') AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0) AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER;"

            End If

            'ASSAYANALYTEKNOWN.ANALYTEFLAGPERCENT

            strSQL = str1 & str2 & str3 & str4
            '''Console.WriteLine("rsFindNomConc: " & strSQL)
            '''debugWriteLine(strSQL)
            ''Console.WriteLine("rsFindNomConc: " & strSQL)

            'must make client side because will use rs.filter later
            rsFindNomConc.CursorLocation = CursorLocationEnum.adUseClient
            rsFindNomConc.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsFindNomConc.ActiveConnection = Nothing

            '****end setup rs

            str1 = "Retrieving Watson Data...15.01 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            If boolANSI Then
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS "
                str2 = "FROM ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) <> 3) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
            Else

            End If

            'new SQL statement incorporating anaylteid
            If boolANSI Then
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<>3) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"
            Else

            End If

            'include runtypeid = 3 : PSAE
            If boolAccess Then
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                '20160221 LEE: added , ANALYTICALRUNSAMPLE.REPLICATENUMBER
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.REPLICATENUMBER "
                str2 = "FROM ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                '20160222 LEE: added DECISIONREASON for Excluded samples
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ANARUNPEAKDECISION.DECISIONREASON "
                str2 = "FROM ANARUNPEAKDECISION RIGHT JOIN (ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (ANARUNPEAKDECISION.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANARUNPEAKDECISION.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANARUNPEAKDECISION.STUDYID = " & strAnaRunPeak & ".STUDYID) "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)>0) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                ''20170928 LEE: added SAMPLETYPEID (matrix) for multiple matrix studies
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ANARUNPEAKDECISION.DECISIONREASON, CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (ANARUNPEAKDECISION RIGHT JOIN (ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) ON (ANARUNPEAKDECISION.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANARUNPEAKDECISION.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANARUNPEAKDECISION.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)>0) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.REPLICATENUMBER, ANARUNPEAKDECISION.DECISIONREASON, CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (ANARUNPEAKDECISION RIGHT JOIN (ASSAYANALYTES INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) ON (ANARUNPEAKDECISION.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANARUNPEAKDECISION.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (ANARUNPEAKDECISION.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)>0) AND ((ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"



            Else

                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                '20160221 LEE: added , ANALYTICALRUNSAMPLE.REPLICATENUMBER
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.REPLICATENUMBER "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                '20160222 LEE: added DECISIONREASON for Excluded samples
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.REPLICATENUMBER, " & strSchema & ".ANARUNPEAKDECISION.DECISIONREASON "
                str2 = "FROM " & strSchema & ".ANARUNPEAKDECISION RIGHT JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) ON (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".ANARUNPEAKDECISION.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANARUNPEAKDECISION.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                '20170928 LEE: added SAMPLETYPEID (matrix) for multiple matrix studies
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.REPLICATENUMBER, " & strSchema & ".ANARUNPEAKDECISION.DECISIONREASON, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (" & strSchema & ".ANARUNPEAKDECISION RIGHT JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ANARUNPEAKDECISION.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNPEAKDECISION.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"

                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, ROUND(" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR," & intDFDec & ") AS ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.REPLICATENUMBER, " & strSchema & ".ANARUNPEAKDECISION.DECISIONREASON, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (" & strSchema & ".ANARUNPEAKDECISION RIGHT JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ANARUNPEAKDECISION.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANARUNPEAKDECISION.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNPEAKDECISION.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)>0) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL) Is Not Null) AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND) Like '%QC%') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER;"


            End If

            'designsampleid
            'nomconc

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblBCQCConcs: " & strSQL)

            If rs3.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs3.Close()
            End If
            rs3.CursorLocation = ADODB.CursorLocationEnum.adUseClient

            Try
                rs3.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                var1 = var1 'debug
            End Try

            var1 = rs3.RecordCount

            rs3.ActiveConnection = Nothing

            str1 = "Retrieving Watson Data...15.02 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            tblBCQCConcs.Clear()
            tblBCQCConcs.AcceptChanges()
            tblBCQCConcs.BeginLoadData()
            daDoPr.Fill(tblBCQCConcs, rs3)
            tblBCQCConcs.EndLoadData()

            'note that tblBCQCConcs has NomConc and QCLABEL and AnalyteDescription columns added in FillDoPrepareTables that gets called earlier

            ''add a column
            'If tblBCQCConcs.Columns.Contains("AnalyteDescription") Then
            'Else
            '    Dim col101 As New DataColumn
            '    col101.ColumnName = "AnalyteDescription"
            '    tblBCQCConcs.Columns.Add(col101)
            'End If

            Call FixTableConcentrations(tblBCQCConcs)


            str1 = "Retrieving Watson Data...15.03 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            Try
                If rs3.EOF And rs3.BOF Then
                Else
                    rs3.MoveFirst()
                End If
            Catch ex As Exception
                var1 = ex.Message
            End Try


            Dim boolMF As Boolean
            boolMF = False
            Count1 = -1


            Do Until rs3.EOF
                Count1 = Count1 + 1
                boolMF = True
                'Dim drow As New DataRow
                'drow = tblBCQCConcs.NewRow
                'For Each fld In rs3.Fields
                '    drow(fld.Name) = fld.Value
                'Next

                'find nomConcentration
                rsFindNomConc.Filter = ""
                'str1 = "ASSAYID = " & tblBCQCConcs.Rows.Item(Count1).Item("ASSAYID") & " AND LEVELNUMBER = " & tblBCQCConcs.Rows.Item(Count1).Item("ASSAYLEVEL") & " AND MASTERASSAYID = " & tblBCQCConcs.Rows.Item(Count1).Item("MASTERASSAYID") & " AND ANALYTEINDEX = " & tblBCQCConcs.Rows.Item(Count1).Item("ANALYTEINDEX") & " AND ANALYTEID = " & tblBCQCConcs.Rows.Item(Count1).Item("ANALYTEID")
                str1 = "ASSAYID = " & tblBCQCConcs.Rows.Item(Count1).Item("ASSAYID") & " AND LEVELNUMBER = " & tblBCQCConcs.Rows.Item(Count1).Item("ASSAYLEVEL") & " AND ANALYTEINDEX = " & tblBCQCConcs.Rows.Item(Count1).Item("ANALYTEINDEX") & " AND ANALYTEID = " & tblBCQCConcs.Rows.Item(Count1).Item("ANALYTEID")
                rsFindNomConc.Filter = str1
                int1 = rsFindNomConc.RecordCount

                If int1 = 0 Then
                    var1 = 0
                    var2 = ""
                Else
                    var1 = rsFindNomConc.Fields("CONCENTRATION").Value
                    var2 = rsFindNomConc.Fields("ID").Value
                    'drow("NomConc") = CDec(NZ(rsFindNomConc.Fields("CONCENTRATION").Value, 0))
                End If
                tblBCQCConcs.Rows.Item(Count1).BeginEdit()
                tblBCQCConcs.Rows.Item(Count1).Item("NomConc") = CDec(NZ(var1, 0))
                tblBCQCConcs.Rows.Item(Count1).Item("QCLABEL") = NZ(var2, "")
                tblBCQCConcs.Rows.Item(Count1).EndEdit()

                rs3.MoveNext()
            Loop
            If boolMF Then
                rs3.MoveFirst()
            End If

            ''debug
            ''Console.WriteLine("Start tblBCQCConcs")
            'var1 = ""
            'For Count1 = 0 To tblBCQCConcs.Columns.Count - 1
            '    var2 = tblBCQCConcs.Columns(Count1).ColumnName
            '    var1 = var1 & ChrW(9) & var2
            'Next
            ''Console.WriteLine(var1)
            'For Count2 = 0 To tblBCQCConcs.Rows.Count - 1
            '    var1 = ""
            '    For Count1 = 0 To tblBCQCConcs.Columns.Count - 1
            '        var2 = tblBCQCConcs.Rows(Count2).Item(Count1)
            '        var1 = var1 & ChrW(9) & var2.ToString
            '    Next
            '    'Console.WriteLine(var1)
            'Next
            ''Console.WriteLine("End tblBCQCConcs")

            str1 = "Retrieving Watson Data...16 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            If ctPB > frmH.pb1.Maximum Then
                ctPB = 1
            End If
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'rs3.ActiveConnection = Nothing

            'begin doing statistics
            If boolANSI Then
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANARUNREGPARAMETERS.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAY.MASTERASSAYID "
                str2 = "FROM ASSAY INNER JOIN ((ANARUNREGPARAMETERS INNER JOIN ANALYTICALRUN ON (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) ON ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<> 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANALYTICALRUN.ASSAYID;"
            Else
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANARUNREGPARAMETERS.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAY.MASTERASSAYID "
                str2 = "FROM ASSAY, ANARUNREGPARAMETERS, ANALYTICALRUN, ASSAYANALYTEKNOWN "
                str2 = str2 & "WHERE (((ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) AND ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "AND (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<> 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANALYTICALRUN.ASSAYID;"
            End If

            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTEKNOWN.LEVELNUMBER, ANARUNREGPARAMETERS.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.CONCENTRATION, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAY.MASTERASSAYID "
                str2 = "FROM ASSAY INNER JOIN ((ANARUNREGPARAMETERS INNER JOIN ANALYTICALRUN ON (ANARUNREGPARAMETERS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNREGPARAMETERS.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON ANALYTICALRUN.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) ON ASSAY.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<> 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANALYTICALRUN.ASSAYID;"
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ASSAY.MASTERASSAYID "
                str2 = "FROM " & strSchema & ".ASSAY INNER JOIN ((" & strSchema & ".ANARUNREGPARAMETERS INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANARUNREGPARAMETERS.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANARUNREGPARAMETERS.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON " & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) ON " & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID "
                str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) Like '%QC%') AND ((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID)<> 3)) "
                str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID;"
            End If
            strSQL = str1 & str2 & str3 & str4
            '''''debugwriteline("rsF: ")
            '''''debugwriteline(strSQL)
            '
            If rsF.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsF.Close()
            End If
            rsF.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rsF.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '''''''''''''''''''''''''''''''''''''Console.WriteLine(strSQL)
            rsF.ActiveConnection = Nothing

            tblQCF.Clear()
            tblQCF.AcceptChanges()
            tblQCF.BeginLoadData()
            daDoPr.Fill(tblQCF, rsF)
            tblQCF.EndLoadData()

            FixTableConcentrations(tblQCF)

            If rsF.EOF And rsF.BOF Then
            Else
                rsF.MoveFirst()
            End If
            'rsF.ActiveConnection = Nothing

            '******

            str1 = "Retrieving Watson Data...17 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'get regression table information for later reporting
            'fill tblregcon
            If boolANSI Then
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID,  ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) "
                str3 = "WHERE (((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<> 3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"
            Else
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID,  ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ASSAY, ANALYTICALRUN, ANALYTICALRUNANALYTES, ANARUNREGPARAMETERS "
                str2 = str2 & "WHERE (((ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID)) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) "
                str3 = "AND (((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID)<> 3) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"
            End If

            'INCLUDE PSAE
            If boolANSI Then
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID,  ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) "
                str3 = "WHERE (((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"
            Else
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID,  ANALYTICALRUN.RUNTYPEID "
                str2 = "FROM ASSAY, ANALYTICALRUN, ANALYTICALRUNANALYTES, ANARUNREGPARAMETERS "
                str2 = str2 & "WHERE (((ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID)) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) "
                str3 = "AND (((ANARUNREGPARAMETERS.STUDYID)=" & wStudyID & ") AND ((ANALYTICALRUN.RUNTYPEID) > 0) AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"
            End If

            'INCLUDE Wting and RegressionText
            If boolANSI Then
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
                str2 = "FROM (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN CONFIGREGRESSIONTYPES ON ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID "
                str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"
            Else
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ASSAY.MASTERASSAYID, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
                str2 = "FROM ASSAY, ANALYTICALRUN, ANALYTICALRUNANALYTES, ANARUNREGPARAMETERS, CONFIGREGRESSIONTYPES "
                str2 = str2 & "WHERE ((((ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX)) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) AND ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID "
                str3 = "AND (((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

            End If

            If boolAccess Then


                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ASSAY.MASTERASSAYID, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
                str2 = "FROM ASSAYANALYTES INNER JOIN ((ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN CONFIGREGRESSIONTYPES ON ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                'Added ASSAY.ASSAYID
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ASSAY.MASTERASSAYID, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ASSAY.ASSAYID "
                str2 = "FROM ASSAYANALYTES INNER JOIN ((ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN CONFIGREGRESSIONTYPES ON ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160212 LEE: Added Matrix
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ASSAY.MASTERASSAYID, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ASSAY.ASSAYID, CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM CONFIGSAMPLETYPES INNER JOIN (ASSAYANALYTES INNER JOIN ((ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN CONFIGREGRESSIONTYPES ON ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) ON CONFIGSAMPLETYPES.SAMPLETYPEKEY = ASSAY.SAMPLETYPEKEY "
                str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160213 LEE: Based new on tblAccAnalRuns
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.RSQUARED, ANARUNREGPARAMETERS.PARAMETERVALUE, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR "
                str2 = "FROM CONFIGREGRESSIONTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON CONFIGREGRESSIONTYPES.REGRESSIONID = ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20161209 LEE:  added , ANALYTICALRUN.RUNSTARTDATE to allow sort Regr param table consistent with Calibr Std table logic
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.RSQUARED, ANARUNREGPARAMETERS.PARAMETERVALUE, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, ANALYTICALRUN.RUNSTARTDATE "
                str2 = "FROM CONFIGREGRESSIONTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON CONFIGREGRESSIONTYPES.REGRESSIONID = ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"


                '20161211 LEE: , ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC  added NM and VEC to allow reporting of LLOQ and ULOQ in Regr Params table
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.RSQUARED, ANARUNREGPARAMETERS.PARAMETERVALUE, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC "
                str2 = "FROM CONFIGREGRESSIONTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON CONFIGREGRESSIONTYPES.REGRESSIONID = ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20170505 LEE:  Added , ASSAYANALYTES.HEIGHTORAREA
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.RSQUARED, ANARUNREGPARAMETERS.PARAMETERVALUE, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, ASSAYANALYTES.HEIGHTORAREA "
                str2 = "FROM CONFIGREGRESSIONTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON CONFIGREGRESSIONTYPES.REGRESSIONID = ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) <> 3) And ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"
                ', ASSAYANALYTES.HEIGHTORAREA

                'matrix
            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.STUDYID, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".CONFIGREGRESSIONTYPES ON " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                str3 = "WHERE(((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                'Added ASSAY.ASSAYID
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.STUDYID, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ASSAY.ASSAYID "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".CONFIGREGRESSIONTYPES ON " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                str3 = "WHERE(((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160212 LEE: Added Matrix
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.STUDYID, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM " & strSchema & ".CONFIGSAMPLETYPES INNER JOIN (" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".CONFIGREGRESSIONTYPES ON " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) ON " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY = " & strSchema & ".ASSAY.SAMPLETYPEKEY "
                str3 = "WHERE(((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3) And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160213 LEE: Based new on tblAccAnalRuns
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR "
                str2 = "FROM " & strSchema & ".CONFIGREGRESSIONTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID = " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) <> 3) And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20161209 LEE:  added , ANALYTICALRUN.RUNSTARTDATE to allow sort Regr param table consistent with Calibr Std table logic
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE "
                str2 = "FROM " & strSchema & ".CONFIGREGRESSIONTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID = " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) <> 3) And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20161211 LEE: , ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC  added NM and VEC to allow reporting of LLOQ and ULOQ in Regr Params table
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC "
                str2 = "FROM " & strSchema & ".CONFIGREGRESSIONTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID = " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) <> 3) And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20170505 LEE:  Added , ASSAYANALYTES.HEIGHTORAREA
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC, " & strSchema & ".ASSAYANALYTES.HEIGHTORAREA "
                str2 = "FROM " & strSchema & ".CONFIGREGRESSIONTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID = " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) <> 3) And ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS) = 3)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"


            End If

            strSQL = str1 & str2 & str3 & str4
            'Console.WriteLine("tblRegCon: " & strSQL)
            If rsF1.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsF1.Close()
            End If
            rsF1.CursorLocation = CursorLocationEnum.adUseClient
            rsF1.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsF1.ActiveConnection = Nothing
            'add columns to table tblregcon
            tblRegCon.Clear()
            tblRegCon.AcceptChanges()
            tblRegCon.BeginLoadData()
            daDoPr.Fill(tblRegCon, rsF1)
            tblRegCon.EndLoadData()

            If rsF1.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsF1.Close()
            End If
            rsF1 = Nothing


            'get from tblRegCon
            Dim tblAAA As DataTable = tblRegCon.DefaultView.ToTable("a", True, "HEIGHTORAREA")
            For Count2 = 0 To tblAAA.Rows.Count - 1
                var1 = NZ(tblAAA.Rows(Count2).Item("HEIGHTORAREA"), 0)
                If var1 = 0 Then
                    str2 = "peak area"
                Else
                    str2 = "peak height"
                End If
                If Count2 = 0 Then
                    str1 = str2
                Else
                    str1 = str1 & ", " & str2
                End If
            Next
            With tblWatsonData
                int1 = FindRow("Integration Type", tblWatsonData, "Item")
                .Rows.Item(int1).Item(1) = str1
            End With
            'Sheets("Data").Range("IntegrationType").Offset(0, 1).Value = str1



            'now do tblRegConAll
            Dim rsF2 As New ADODB.Recordset
            If boolAccess Then

                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ASSAY.MASTERASSAYID, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
                str2 = "FROM ASSAYANALYTES INNER JOIN ((ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID)) INNER JOIN CONFIGREGRESSIONTYPES ON ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID) ON (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160212 LEE: Added matrix
                str1 = "SELECT DISTINCT ANARUNREGPARAMETERS.STUDYID, ANARUNREGPARAMETERS.RUNID, ASSAY.MASTERASSAYID, ANARUNREGPARAMETERS.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, ANARUNREGPARAMETERS.PARAMETERVALUE, ANALYTICALRUNANALYTES.RSQUARED, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (ASSAYANALYTES INNER JOIN ((ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (ANALYTICALRUN.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN CONFIGREGRESSIONTYPES ON ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE(((ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANARUNREGPARAMETERS.RUNID, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160213 LEE: Based new on tblAccAnalRuns
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.RSQUARED, ANARUNREGPARAMETERS.PARAMETERVALUE, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR "
                str2 = "FROM CONFIGREGRESSIONTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON CONFIGREGRESSIONTYPES.REGRESSIONID = ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20161209 LEE:  added , ANALYTICALRUN.RUNSTARTDATE to allow sort Regr param table consistent with Calibr Std table logic
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.RSQUARED, ANARUNREGPARAMETERS.PARAMETERVALUE, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, ANALYTICALRUN.RUNSTARTDATE "
                str2 = "FROM CONFIGREGRESSIONTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON CONFIGREGRESSIONTYPES.REGRESSIONID = ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20161211 LEE: , ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC  added NM and VEC to allow reporting of LLOQ and ULOQ in Regr Params table
                str1 = "SELECT DISTINCT ANALYTICALRUN.STUDYID, ASSAY.MASTERASSAYID, ASSAY.ASSAYID, ANALYTICALRUN.RUNID, ANALYTICALRUN.RUNTYPEID, ASSAYANALYTES.ANALYTEID, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.RSQUARED, ANARUNREGPARAMETERS.PARAMETERVALUE, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, CONFIGREGRESSIONTYPES.REGRESSIONTEXT, ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, ANALYTICALRUN.RUNSTARTDATE, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC "
                str2 = "FROM CONFIGREGRESSIONTYPES INNER JOIN (((GLOBALANALYTES INNER JOIN (((ANALYTICALRUN INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ANALYTICALRUNANALYTES ON (ANALYTICALRUN.RUNID = ANALYTICALRUNANALYTES.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON GLOBALANALYTES.GLOBALANALYTEID = ASSAYANALYTES.ANALYTEID) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANARUNREGPARAMETERS ON (ANALYTICALRUNANALYTES.RUNID = ANARUNREGPARAMETERS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNREGPARAMETERS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNREGPARAMETERS.STUDYID)) ON CONFIGREGRESSIONTYPES.REGRESSIONID = ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY ANALYTICALRUN.RUNID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNANALYTES.ANALYTEINDEX, ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"


            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.STUDYID, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
                str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID)) ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID)) INNER JOIN " & strSchema & ".CONFIGREGRESSIONTYPES ON " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID) ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) "
                str3 = "WHERE(((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160212 LEE: Added matrix
                str1 = "SELECT DISTINCT " & strSchema & ".ANARUNREGPARAMETERS.STUDYID, " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID "
                str2 = "FROM (" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID =" & strSchema & ". ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ASSAY.STUDYID =" & strSchema & ". ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".CONFIGREGRESSIONTYPES ON " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER = " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                str3 = "WHERE(((" & strSchema & ".ANARUNREGPARAMETERS.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANARUNREGPARAMETERS.RUNID, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20160213 LEE: Based new on tblAccAnalRuns
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR "
                str2 = "FROM " & strSchema & ".CONFIGREGRESSIONTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID = " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20161209 LEE:  added , ANALYTICALRUN.RUNSTARTDATE to allow sort Regr param table consistent with Calibr Std table logic
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE "
                str2 = "FROM " & strSchema & ".CONFIGREGRESSIONTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID = " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"

                '20161211 LEE: , ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC  added NM and VEC to allow reporting of LLOQ and ULOQ in Regr Params table
                str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUN.STUDYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAY.ASSAYID, " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.RSQUARED, " & strSchema & ".ANARUNREGPARAMETERS.PARAMETERVALUE, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID, " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONTEXT, " & strSchema & ".ANALYTICALRUNANALYTES.WEIGHTINGFACTOR, " & strSchema & ".ANALYTICALRUN.RUNSTARTDATE, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC "
                str2 = "FROM " & strSchema & ".CONFIGREGRESSIONTYPES INNER JOIN (((" & strSchema & ".GLOBALANALYTES INNER JOIN (((" & strSchema & ".ANALYTICALRUN INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ANALYTICALRUNANALYTES ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNANALYTES.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANARUNREGPARAMETERS ON (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNREGPARAMETERS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNREGPARAMETERS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNREGPARAMETERS.STUDYID)) ON " & strSchema & ".CONFIGREGRESSIONTYPES.REGRESSIONID = " & strSchema & ".ANALYTICALRUNANALYTES.REGRESSIONIDENTIFIER "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".ANALYTICALRUN.RUNTYPEID) > 0)) "
                str4 = "ORDER BY " & strSchema & ".ANALYTICALRUN.RUNID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX, " & strSchema & ".ANARUNREGPARAMETERS.REGRESSIONPARAMETERID;"


            End If

            strSQL = str1 & str2 & str3 & str4
            '''Console.WriteLine("tblRegConAll: " & strSQL)
            If rsF2.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsF2.Close()
            End If
            rsF2.CursorLocation = CursorLocationEnum.adUseClient
            rsF2.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsF2.ActiveConnection = Nothing
            'add columns to table tblregconall
            tblRegConAll.Clear()
            tblRegConAll.AcceptChanges()
            tblRegConAll.BeginLoadData()
            daDoPr.Fill(tblRegConAll, rsF2)
            tblRegConAll.EndLoadData()

            If rsF2.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsF2.Close()
            End If
            rsF2 = Nothing

            '******

            If boolAccess Then
                str1 = "SELECT ANARUNANALYTERESULTS.STUDYID, SAMPRESCONFLICTDEC.ANALYTEID, SAMPRESCONFLICTDEC.DESIGNSAMPLEID, ANARUNANALYTERESULTS.RUNID, ANARUNANALYTERESULTS.CONCENTRATION "
                str2 = "FROM SAMPRESCONFLICTDEC INNER JOIN ANARUNANALYTERESULTS ON (SAMPRESCONFLICTDEC.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (SAMPRESCONFLICTDEC.CONCENTRATION = ANARUNANALYTERESULTS.CONCENTRATION) "
                str3 = "WHERE(((ANARUNANALYTERESULTS.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY SAMPRESCONFLICTDEC.ANALYTEID, SAMPRESCONFLICTDEC.DESIGNSAMPLEID, ANARUNANALYTERESULTS.RUNID;"
            Else
                str1 = "SELECT " & strSchema & ".ANARUNANALYTERESULTS.STUDYID, " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID, " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION "
                str2 = "FROM " & strSchema & ".SAMPRESCONFLICTDEC INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTDEC.CONCENTRATION = " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION) "
                str3 = "WHERE(((" & strSchema & ".ANARUNANALYTERESULTS.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID, " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.RUNID;"
            End If
            '

            strSQL = str1 & str2 & str3 & str4

            '''Console.WriteLine("tblGetDecRunID:  " & strSQL)
            Dim rs100 As New ADODB.Recordset
            If rs100.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs100.Close()
            End If
            rs100.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs100.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '''''''''''''''''''''''''''''''''''''Console.WriteLine(strSQL)
            rs100.ActiveConnection = Nothing

            tblGetDecRunID.Clear()
            tblGetDecRunID.AcceptChanges()
            tblGetDecRunID.BeginLoadData()
            daDoPr.Fill(tblGetDecRunID, rs100)
            tblGetDecRunID.EndLoadData()

            If rs100.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs100.Close()
            End If
            rs100 = Nothing

            '******

            '******

            If boolAccess Then
                str1 = "SELECT DISTINCT SAMPRESCONFLICTDEC.STUDYID, SAMPRESCONFLICTDEC.ANALYTEID, SAMPRESCONFLICTDEC.DESIGNSAMPLEID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.CONCENTRATION "
                str2 = "FROM SAMPRESCONFLICTDEC "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY SAMPRESCONFLICTDEC.RECORDTIMESTAMP DESC;"

                '20160313 LEE: Added RUNID
                str1 = "SELECT DISTINCT SAMPRESCONFLICTDEC.STUDYID, SAMPRESCONFLICTDEC.ANALYTEID, SAMPRESCONFLICTDEC.DESIGNSAMPLEID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTDEC.CONCENTRATION, SAMPRESCONFLICTDEC.RUNID "
                str2 = "FROM SAMPRESCONFLICTDEC "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY SAMPRESCONFLICTDEC.RECORDTIMESTAMP DESC;"

            Else
                str1 = "SELECT DISTINCT " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID, " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID, " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".SAMPRESCONFLICTDEC.CONCENTRATION "
                str2 = "FROM " & strSchema & ".SAMPRESCONFLICTDEC "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP DESC;"

                '20160313 LEE: Added RUNID
                str1 = "SELECT DISTINCT " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID, " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID, " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".SAMPRESCONFLICTDEC.CONCENTRATION, " & strSchema & ".SAMPRESCONFLICTDEC.RUNID "
                str2 = "FROM " & strSchema & ".SAMPRESCONFLICTDEC "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP DESC;"

            End If
            '

            strSQL = str1 & str2 & str3

            '''Console.WriteLine("tblGetDecRunID:  " & strSQL)
            Dim rs101 As New ADODB.Recordset
            If rs101.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs101.Close()
            End If
            rs101.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rs101.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '''''''''''''''''''''''''''''''''''''Console.WriteLine(strSQL)
            rs101.ActiveConnection = Nothing

            tblSAMPRESCONFLICTDEC.Clear()
            tblSAMPRESCONFLICTDEC.AcceptChanges()
            tblSAMPRESCONFLICTDEC.BeginLoadData()
            daDoPr.Fill(tblSAMPRESCONFLICTDEC, rs101)
            tblSAMPRESCONFLICTDEC.EndLoadData()

            If rs101.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs101.Close()
            End If
            rs101 = Nothing

            '******


            str1 = "Retrieving Watson Data...18 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'determine # of accepted analytical runs for each analyte
            For Count1 = 1 To ctAnalytes
                str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1)
                Erase drows
                drows = tblRegCon.Select(str1)
                int1 = drows.Length
                int2 = 0
                For Count2 = 0 To int1 - 1
                    var1 = drows(Count2).Item("RUNANALYTEREGRESSIONSTATUS")
                    If var1 = 3 Then
                        int2 = int2 + 1
                    End If
                Next

                'determine number of regression constants
                Dim dvAA As System.Data.DataView
                Dim tblAA As System.Data.DataTable
                dvAA = New DataView(tblRegCon)
                tblAA = dvAA.ToTable("a", True, "REGRESSIONPARAMETERID") 'returns distinct regression parameters
                int1 = tblAA.Rows.Count
                arrAnalytes(7, Count1) = int2 / int1 '# of accepted runs. Divide by int1 because there are int1 records
                '(parametervalues for A, B, C, etc) for each accepted run
            Next
            '1=RUNID, 2=AnalyteIndex, 3=REGRESSIONPARAMETERID(1=Slope, 2=YInt, 3=R2),4=PARAMETERVALUE
            '1=RUNID,  2=Slope, 3=YInt, 4=R2

            '******

            'int10 = FindRow("Calibration Levels", dgWatsonAnalRef)
            'int20 = FindRow("Minimum r^2", dgWatsonAnalRef)
            int30 = FindRow("QC Mean Accuracy Min", tblWatsonAnalRefTable, "Item")
            int40 = FindRow("QC Mean Accuracy Max", tblWatsonAnalRefTable, "Item")
            int50 = FindRow("QC Precision Min", tblWatsonAnalRefTable, "Item")
            int60 = FindRow("QC Precision Max", tblWatsonAnalRefTable, "Item")

            Dim ctQCs As Short
            'problem is here
            For Count1 = 1 To ctAnalytes

                '***here
                Dim arrBCQCs(5, 200) '1=LevelNumber, 2=Concentration, 3=ID, 4=#ofReplicates
                Dim arrBCQCConcs(7, 100) '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=AssayID
                Dim arrQCAI(3, 200) '1=arrBCQCs number, 2=AssayID,3=NomConcentration
                '(1=LevelNumber, 2=Concentration, 3=RunID) <--???

                'find number of QC levels
                Count2 = 0
                Count3 = 1
                rs4.Filter = ""
                'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1)
                str1 = "ANALYTEID = " & arrAnalytes(2, Count1)
                var1 = arrAnalytes(12, Count1)
                var2 = arrAnalytes(3, Count1)
                rs4.Filter = str1
                int1 = rs4.RecordCount
                numQCLevels = 0
                Do Until rs4.EOF
                    'before recording, ensure that qc level was ever used
                    var3 = rs4.Fields("LevelNumber").Value
                    var4 = rs4.Fields("Concentration").Value
                    var5 = rs4.Fields("ANALYTEID").Value
                    'str1 = "ASSAYLEVEL = " & var3 & " AND ANALYTEINDEX = " & var2 & " and MASTERASSAYID = " & var1
                    'str1 = "ASSAYLEVEL = " & var3 & " AND ANALYTEINDEX = " & var2 & " and MASTERASSAYID = " & var1 & " AND NOMCONC = " & var4
                    'str1 = "ASSAYLEVEL = " & var3 & " AND ANALYTEINDEX = " & var2 & " and MASTERASSAYID = " & var1 & " AND NOMCONC = " & var4 & " AND ANALYTEID = " & var5
                    str1 = "ASSAYLEVEL = " & var3 & " AND NOMCONC = " & var4 & " AND ANALYTEID = " & var5
                    Erase drows
                    drows = tblBCQCConcs.Select(str1)
                    int1 = drows.Length
                    If int1 = 0 Then
                    Else 'continue
                        'ensure assayid was used

                        Count2 = Count2 + 1
                        If Count2 > UBound(arrBCQCs, 2) Then
                            ReDim Preserve arrBCQCs(5, UBound(arrBCQCs, 2) + 100)
                        End If
                        arrBCQCs(1, Count2) = rs4.Fields("LevelNumber").Value
                        var6 = rs4.Fields("CONCENTRATION").Value 'test
                        arrBCQCs(2, Count2) = rs4.Fields("CONCENTRATION").Value
                        arrBCQCs(3, Count2) = rs4.Fields("ID").Value
                    End If
                    var4 = rs4.Fields("ID").Value
                    'If InStr(var4, "Dil", CompareMethod.Text) > 0 Then
                    'Else
                    '    numQCLevels = numQCLevels + 1
                    'End If
                    numQCLevels = numQCLevels + 1
                    rs4.MoveNext()
                Loop
                var1 = arrBCQCs(1, Count2)
                ctQCs = Count2

                'record items in tblWatsonAnalRef
                int1 = FindRow("# of QC Levels", tblWatsonAnalRefTable, "Item")

                tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = numQCLevels

                'fill in arrctanalytes
                arrctQCs(1, Count1) = var2 'analyte id
                arrctQCs(2, Count1) = var1 'max Level Number
                arrctQCs(3, Count1) = ctQCs 'number of qcs

                'rs.Close()

                'determine appropriate assayid's
                '1=arrBCQCs number, 2=AssayID,3=NomConcentration
                Count3 = 0

                'rs1.Filter = ""
                'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1)
                'rs1.Filter = str1
                For Count2 = 1 To ctQCs
                    rs1.Filter = ""
                    str3 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND CONCENTRATION = " & arrBCQCs(2, Count2) & " and MASTERASSAYID = " & arrAnalytes(12, Count1)
                    rs1.Filter = str3
                    Do Until rs1.EOF
                        Count3 = Count3 + 1
                        If Count3 > UBound(arrQCAI, 2) Then
                            ReDim Preserve arrQCAI(3, UBound(arrQCAI, 2) + 100)
                        End If
                        var1 = rs1.Fields("AssayID").Value
                        arrQCAI(1, Count3) = Count2
                        arrQCAI(2, Count3) = var1
                        arrQCAI(3, Count3) = rs1.Fields("Concentration").Value
                        ''''''debugPrint Count2 & ", " & var1 & ", " & rs.Fields("Concentration").Value
                        rs1.MoveNext()
                    Loop
                Next
                ctQCAI = Count3

                str1 = "Retrieving Watson Data...19 " & ctPB
                str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
                frmH.lblProgress.Text = str1
                ctPB = ctPB + 1
                If ctPB > frmH.pb1.Maximum Then
                    ctPB = 1
                End If

                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                frmH.lblProgress.Refresh()
                System.Windows.Forms.Application.DoEvents()

                'find #ofReplicates for each level
                numRepDilnQC = 0
                numRepQC = 0
                maxRep = 0
                For Count2 = 1 To ctQCs
                    var1 = arrBCQCs(1, Count2)
                    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND LEVELNUMBER = " & var1 & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " and CONCENTRATION = " & arrBCQCs(2, Count2)
                    ''''''debugwriteline(str1)
                    ''
                    var2 = arrAnalytes(3, Count1)
                    var3 = arrAnalytes(12, Count1)
                    rs20.Filter = ""
                    rs20.Filter = str1
                    var6 = rs20.RecordCount
                    If var6 = 0 Then
                    Else
                        rs20.MoveLast()
                        var4 = rs20.Fields("REPLICATENUMBER").Value
                        arrBCQCs(4, Count2) = rs20.Fields("REPLICATENUMBER").Value
                        If rs20.Fields("REPLICATENUMBER").Value > maxRep Then
                            maxRep = maxRep + 1
                        End If
                        'var5 = arrBCQCs(3, Count2) '????
                    End If

                Next

                'QC Results
                'ReDim arrBCQCConcs(7, 100)
                '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=AssayID
                'find Interpolated QC Standard Concentrations for all analytical runs
                Count2 = 0
                rs3.Filter = ""
                str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1)
                '''''debugwriteline(str1)
                '
                rs3.Filter = str1
                'rs3.MoveFirst()

                If rs3.RecordCount = 0 Then
                Else
                    rs3.MoveFirst()
                    'Info: arrBCQCs():'1=LevelNumber, 2=Concentration, 3=RunID, 4=#ofReplicate
                    Do Until rs3.EOF
                        Count2 = Count2 + 1
                        If Count2 > UBound(arrBCQCConcs, 2) Then
                            ReDim Preserve arrBCQCConcs(7, UBound(arrBCQCConcs, 2) + 100)
                        End If
                        '1=LevelNumber, 2=Concentration, 3=ID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=AssayID
                        'arrBCQCConcs(1, Count2) = Count3
                        arrBCQCConcs(1, Count2) = rs3.Fields("ASSAYLEVEL").Value
                        str2 = NZ(rs3.Fields("SAMPLENAME").Value, "")
                        num1 = NZ(rs3.Fields("CONCENTRATION").Value, 0)
                        'num1 = SigFigOrDec(num1, LSigFig, True)
                        arrBCQCConcs(2, Count2) = num1
                        arrBCQCConcs(3, Count2) = rs3.Fields("RUNID").Value
                        arrBCQCConcs(4, Count2) = rs3.Fields("ELIMINATEDFLAG").Value
                        arrBCQCConcs(5, Count2) = str2
                        arrBCQCConcs(6, Count2) = CDbl(rs3.Fields("ALIQUOTFACTOR").Value)
                        arrBCQCConcs(7, Count2) = rs3.Fields("ASSAYID").Value
                        ''''''debugwriteline("Level: " & arrBCQCConcs(1, Count2) & ", Assayid: " & arrBCQCConcs(7, Count2))
                        ''

                        rs3.MoveNext()
                    Loop
                End If

                '
                rs3.Filter = ""

                str1 = "Retrieving Watson Data...20 " & ctPB
                str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
                frmH.lblProgress.Text = str1
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                frmH.lblProgress.Refresh()
                System.Windows.Forms.Application.DoEvents()


                int1 = Count2
                'int1 = arrAnalytes(7, Count1) '#accepted runs
                inttemprows = int1

                'begin doing statistics
                ReDim arrBCStdActual(inttemprows * 2)
                Dim arrPrec(ctQCs)
                Dim arrAcc(ctQCs)
                ''''''debugwriteline("Start;")
                For Count3 = 1 To ctQCs
                    'int1 = 0
                    'var1 = arrBCQCs(1, Count3) 'level
                    'var2 = arrBCQCs(2, Count3) 'concentration
                    'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND CONCENTRATION = " & var2 & " and MASTERASSAYID = " & arrAnalytes(12, Count1)
                    'rsF.Filter = ""
                    ''open filtered recordset
                    'rsF.Filter = str1

                    '''''''''''''''''''''''''Console.WriteLine(" ")
                    '''''''''''''''''''''''''Console.WriteLine("Begin Assay Level " & var1)
                    'ReDim arrBCStdActual(inttemprows * 2)
                    'Do Until rsF.EOF

                    '    var3 = rsF.Fields("ASSAYID").Value
                    '    'arrBCQCConcs:'1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=AssayID
                    '    For Count5 = 1 To inttemprows
                    '        var4 = arrBCQCConcs(7, Count5) 'ASSAYID
                    '        var5 = arrBCQCConcs(1, Count5) 'LEVEL
                    '        If var3 = var4 And var1 = var5 Then
                    '            var6 = arrBCQCConcs(4, Count5) 'Eliminated Flag
                    '            If StrComp(var6, "Y", vbTextCompare) = 0 Or IsDBNull(arrBCQCConcs(2, Count5)) Then
                    '            Else 'this section takes care of aliquot factor
                    '                int1 = int1 + 1
                    '                If int1 > UBound(arrBCStdActual) Then
                    '                    ReDim Preserve arrBCStdActual(UBound(arrBCStdActual) + 50)
                    '                End If
                    '                arrBCStdActual(int1) = CDbl(SigFigOrDec(CDbl(arrBCQCConcs(2, Count5) / arrBCQCConcs(6, Count5)), LSigFig, True))
                    '                var7 = arrBCStdActual(int1)
                    '                ''''''''''''''''''''''''Console.WriteLine(arrBCStdActual(int1))
                    '            End If
                    '        End If

                    '    Next
                    '    rsF.MoveNext()
                    'Loop
                    '''''''''''''''''''''''''Console.WriteLine("End Assay Level " & var1)
                    '''''''''''''''''''''''''Console.WriteLine(" ")


                    '*****start from style
                    int1 = 0
                    var2 = arrBCQCs(1, Count3) '.Item("LevelNumber")
                    var3 = arrBCQCs(2, Count3) 'CONCENTRATION
                    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3
                    'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ASSAYLEVEL = " & Count3
                    Erase drows
                    drows = tblBCQCConcs.Select(str1)
                    int2 = drows.Length
                    ReDim arrBCStdActual(int2) 're-use this array
                    For Count5 = 0 To int2 - 1
                        'num1 = NZ(drows(Count3).Item("CONCENTRATION"), 0)
                        'num1 = SigFigOrDec(CDec(num1), LSigFig, True)
                        'frmh.arrBCStdConcs(2, Count3 + 1) = num1
                        'frmh.arrBCStdConcs(3, Count3 + 1) = drows(Count3).Item("RUNID")
                        'frmh.arrBCStdConcs(4, Count3 + 1) = drows(Count3).Item("ELIMINATEDFLAG")
                        var1 = drows(Count5).Item("ELIMINATEDFLAG") 'frmh.arrBCStdConcs(4, Count5)
                        If StrComp(var1, "Y", vbTextCompare) = 0 Or IsDBNull(drows(Count5).Item("CONCENTRATION")) Then 'exclude value
                        Else
                            int1 = int1 + 1
                            num1 = NZ(drows(Count5).Item("CONCENTRATION"), 0)
                            num2 = NZ(drows(Count5).Item("ALIQUOTFACTOR"), 1)
                            num3 = num1 / num2
                            num1 = SigFigOrDec(CDec(num3), LSigFig, False)
                            arrBCStdActual(int1) = num1
                            'var7 = frmh.arrBCStdConcs(2, Count5)
                        End If
                    Next
                    'determine Sum
                    numSum = 0
                    numMean = SigFigOrDec(Mean(int1, arrBCStdActual), LSigFig, False)
                    numSD = SigFigOrDec(StdDev(int1, arrBCStdActual), LSigFig, False)
                    '***End BCStds

                    '*****end from style

                    rsF.Filter = ""
                    If int1 = 0 Then
                    Else
                        numMean = SigFigOrDec(Mean(int1, arrBCStdActual), LSigFig, False)
                        numSD = SigFigOrDec(StdDev(int1, arrBCStdActual), LSigFig, False)

                        If numMean <= 0 Then
                            var2 = 1
                        Else
                            var2 = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(numSD / numMean * 100, intQCDec + 1), intQCDec), strQCDec)) 'for testing
                        End If
                        arrPrec(Count3) = var2 'CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(numSD / numMean * 100, intQCDec + 1), intQCDec), strQCDec))

                        var1 = arrBCQCs(2, Count3)
                        If var1 <= 0 Then
                            var2 = 0
                        Else
                            var2 = CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(((numMean / arrBCQCs(2, Count3)) - 1) * 100, intQCDec + 1), intQCDec), strQCDec))
                        End If
                        arrAcc(Count3) = var2 ' CDec(Format(RoundToDecimalRAFZ(RoundToDecimalRAFZ(((numMean / arrBCQCs(2, Count3)) - 1) * 100, intQCDec + 1), intQCDec), strQCDec))
                        ''''''debugwriteline(CDec(Format(RoundToDecimal(numSD / numMean * 100, 10), "0.0")) & ";")
                    End If

                Next
                ''''''debugwriteline(arrAnalytes(1, Count1) & ";")


                '*****

                'do this after getting tblRegCon
                'get min regression r2
                'int2 = arrAnalytes(3, Count1) 'analyteindex
                Count2 = 0
                '1=RUNID, 2=AnalyteIndex, 3=REGRESSIONPARAMETERID(1=Slope, 2=YInt, 3=R2),4=PARAMETERVALUE
                '1=RUNID,  2=Slope, 3=YInt, 4=R2
                rsAAR.Filter = ""
                'str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and MASTERASSAYID = " & arrAnalytes(12, Count1)
                str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and ANALYTEID = " & arrAnalytes(2, Count1)
                str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " and ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                Dim rowsAAR() As DataRow = tblRegCon.Select(str1)
                For Count2 = 1 To rowsAAR.Length
                    If Count2 > UBound(arrRegCon) Then
                        ReDim Preserve arrRegCon(UBound(arrRegCon) + 50)
                    End If
                    arrRegCon(Count2) = rowsAAR(Count2 - 1).Item("RSQUARED") ' rsAAR.Fields("RSQUARED").Value
                Next
                'Do Until rsAAR.EOF
                '    Count2 = Count2 + 1
                '    If Count2 > UBound(arrRegCon) Then
                '        ReDim Preserve arrRegCon(UBound(arrRegCon) + 50)
                '    End If
                '    arrRegCon(Count2) = rsAAR.Fields("RSQUARED").Value
                '    rsAAR.MoveNext()
                'Loop
                'record R_2
                'str1 = "0."
                'For Count2 = 1 To LR2SigFigs
                '    str1 = str1 & "0"
                'Next

                var3 = GetMin(arrRegCon, Count2)
                var2 = SigFigOrDecString(var3, LR2SigFigs, False)
                str1 = GetRegrDecStr(LR2SigFigs)
                var1 = Format(CDec(var2), str1)
                tblWatsonAnalRefTable.Rows.Item(int20).Item(Count1) = var1



                '*****

                var1 = Format(CDec(GetMin(arrAcc, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int30).Item(Count1) = var1
                var1 = Format(CDec(GetMax(arrAcc, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int40).Item(Count1) = var1
                var1 = Format(CDec(GetMin(arrPrec, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int50).Item(Count1) = var1
                var1 = Format(CDec(GetMax(arrPrec, ctQCs)), strQCDec)
                tblWatsonAnalRefTable.Rows.Item(int60).Item(Count1) = var1
                '******
                'Find QC and Diln QC Reps

                numRepDilnQC = 0
                numRepQC = 0

                str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND ANALYTEID = " & arrAnalytes(2, Count1)
                drows = tblRegCon.Select(str1)
                int10 = drows.Length

                str1 = "Retrieving Watson Data...21 " & ctPB
                str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
                frmH.lblProgress.Text = str1
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                frmH.lblProgress.Refresh()
                System.Windows.Forms.Application.DoEvents()

                'for each accepted analytical run
                Count5 = 0
                Dim int201 As Short
                Dim drowsF() As DataRow
                Dim intF As Short
                'the following is wrong, but gets corrected in sub AssessQCs
                For Count2 = 0 To int10 - 1 Step 2 'step by two because tblRegCon has doublerow entries
                    'need maxRep rows for each accepted run
                    int201 = CInt(drows(Count2).Item("RUNID"))
                    For Count3 = 0 To maxRep - 1
                        'establish array going across table ctQC number of times
                        For Count4 = 1 To ctQCs
                            var2 = arrBCQCs(1, Count4) '.Item("LevelNumber")

                            var3 = arrBCQCs(2, Count4) 'CONCENTRATION

                            Count5 = Count5 + 1
                            str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNID = " & int201 & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3
                            '''''''''''''''''''''''''''''''''''''''Console.WriteLine(id_tblStudies & ": " & str1)
                            Erase drowsF
                            'drowsF = tblBCQCConcs.Select(str1, "RUNSAMPLESEQUENCENUMBER ASC")
                            drowsF = tblBCQCConcs.Select(str1, "RUNSAMPLEORDERNUMBER ASC")
                            intF = drowsF.Length
                            If intF < Count3 + 1 Then 'enter base values
                            Else 'enter in total set of values
                                ''evaluate QC Rep #
                                var8 = drowsF(Count3).Item("AliquotFactor")
                                If var8 = 1 Then
                                    numRepQC = Count3 + 1
                                Else
                                    numRepDilnQC = Count3 + 1
                                End If

                            End If
                        Next
                    Next
                Next

                'record #ofReplicates in tblWatsonAnalRef: QCConcentrationCount
                int1 = FindRow("# of QC Replicates", tblWatsonAnalRefTable, "Item")
                tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = numRepQC
                int1 = FindRow("# of Dilution QC Replicates", tblWatsonAnalRefTable, "Item")
                tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = numRepDilnQC

                'End Find QC and Diln QC Reps
                '*******

                'fill tblQCStds
                'col1.ColumnName = "AnalyteDescription"
                'col2.ColumnName = "LevelNumber"
                'col3.ColumnName = "Concentration"
                'col4.ColumnName = "NumReps"
                'col5.ColumnName = "MasterAssayID"
                'col6.ColumnName = "ID"
                'col7.ColumnName = "Index"
                'col8.ColumnName = "FlagPercent"

                For Count2 = 1 To ctQCs

                    var1 = arrAnalytes(1, Count1) 'Analyte Description
                    var2 = arrBCQCs(1, Count2) 'Level Number
                    var3 = arrBCQCs(2, Count2) 'concentration
                    var4 = arrBCQCs(3, Count2) 'QC ID
                    var6 = arrAnalytes(12, Count1) 'masterassayid
                    var7 = arrAnalytes(2, Count1) 'Analyte ID
                    var8 = arrAnalytes(3, Count1) 'analyteindex
                    str1 = "ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND LEVELNUMBER = " & var2 & " and MASTERASSAYID = " & arrAnalytes(12, Count1) & " and CONCENTRATION = " & arrBCQCs(2, Count2)

                    rs20.Filter = ""
                    rs20.Filter = str1
                    If rs20.RecordCount = 0 Then
                    Else
                        rs20.MoveLast()
                        var5 = rs20.Fields("REPLICATENUMBER").Value

                        'find FlagPercent
                        rsFindNomConc.Filter = ""
                        str1 = "ASSAYID = " & tblBCQCConcs.Rows.Item(Count1).Item("ASSAYID") & " AND LEVELNUMBER = " & tblBCQCConcs.Rows.Item(Count1).Item("ASSAYLEVEL") & " AND MASTERASSAYID = " & tblBCQCConcs.Rows.Item(Count1).Item("MASTERASSAYID") & " AND ANALYTEINDEX = " & tblBCQCConcs.Rows.Item(Count1).Item("ANALYTEINDEX")
                        rsFindNomConc.Filter = str1
                        int1 = rsFindNomConc.RecordCount
                        If int1 = 0 Then
                            var9 = ""
                        Else
                            'var9 = NZ(rsFindNomConc.Fields("FLAGPERCENT").Value, "")
                            var9 = NZ(rsFindNomConc.Fields("ANALYTEFLAGPERCENT").Value, "")
                        End If

                        Dim row As DataRow = tblQCStds.NewRow
                        row.Item("AnalyteDescription") = NZ(var1, "")
                        row.Item("LevelNumber") = CInt(NZ(var2, 1000))
                        row.Item("Concentration") = CDec(NZ(var3, 0))
                        row.Item("QCName") = NZ(var4, "")
                        row.Item("NumReps") = CInt(NZ(var5, 1000))
                        row.Item("MasterAssayID") = CLng(NZ(var6, 1000))
                        row.Item("ID") = CLng(NZ(var7, 1000))
                        row.Item("Index") = CInt(NZ(var8, 1000))
                        row.Item("FlagPercent") = NZ(var9, "")
                        tblQCStds.Rows.Add(row)
                    End If

                Next
            Next

            rsFindNomConc.Close()

            'add AnalyteDescription data to tblbcstds
            For Count1 = 1 To ctAnalytes
                var1 = arrAnalytes(1, Count1) 'analytedescription
                var2 = arrAnalytes(3, Count1) 'analyteindex
                var3 = arrAnalytes(12, Count1) 'masterassayid
                var6 = arrAnalytes(2, Count1) 'analyteid
                For Count2 = 0 To tblBCStds.Rows.Count - 1
                    var4 = tblBCStds.Rows.Item(Count2).Item("ANALYTEINDEX")
                    var5 = tblBCStds.Rows.Item(Count2).Item("MASTERASSAYID")
                    var7 = tblBCStds.Rows.Item(Count2).Item("ANALYTEID")
                    If var2 = var4 And var3 = var5 And var6 = var7 Then
                        tblBCStds.Rows.Item(Count2).BeginEdit()
                        tblBCStds.Rows.Item(Count2).Item("AnalyteDescription") = var1
                        tblBCStds.Rows.Item(Count2).EndEdit()
                    End If
                Next
            Next

            'add AssayID to tblBCStds

            For Count2 = 0 To tblBCStds.Rows.Count - 1
                var4 = tblBCStds.Rows.Item(Count2).Item("ANALYTEINDEX")
                var5 = tblBCStds.Rows.Item(Count2).Item("MASTERASSAYID")
                var7 = tblBCStds.Rows.Item(Count2).Item("ANALYTEID")

            Next

            'add AnalyteDescription data to tblbcqcconcs
            For Count1 = 1 To ctAnalytes
                var1 = arrAnalytes(1, Count1) 'analytedescription
                var2 = arrAnalytes(3, Count1) 'analyteindex
                var3 = arrAnalytes(12, Count1) 'masterassayid
                var6 = arrAnalytes(2, Count1) 'analyte id
                For Count2 = 0 To tblBCQCConcs.Rows.Count - 1
                    var4 = tblBCQCConcs.Rows.Item(Count2).Item("ANALYTEINDEX")
                    var5 = tblBCQCConcs.Rows.Item(Count2).Item("MASTERASSAYID")
                    If var2 = var4 And var3 = var5 Then
                        tblBCQCConcs.Rows.Item(Count2).BeginEdit()
                        tblBCQCConcs.Rows.Item(Count2).Item("AnalyteDescription") = var1
                        tblBCQCConcs.Rows.Item(Count2).EndEdit()
                    End If
                Next
            Next

            '***Start here 5
            'retrieve Calibration Curve Info
            str1 = "Retrieving calibration curve info..."

            'retrieve regression type and weighting factor
            'weighting factor: 0 = 1, 1 = 1/x, 2 = 1/x**2, 3 = 1/y, 4 = 1/y**2

            'THE NEXT IS BAD. WILL RETURN INCORRECT INFORMATION
            'GET FROM TBLREGCON INSTEAD
            'If boolANSI Then
            '    str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.WEIGHTINGFACTOR, ASSAYANALYTES.REGRESSIONIDENTIFIER, CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
            '    str2 = "FROM ASSAYANALYTES INNER JOIN CONFIGREGRESSIONTYPES ON ASSAYANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID "
            '    str3 = "WHERE (((ASSAYANALYTES.STUDYID)=" & wStudyID & "));"
            'Else
            '    str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEINDEX, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.WEIGHTINGFACTOR, ASSAYANALYTES.REGRESSIONIDENTIFIER, CONFIGREGRESSIONTYPES.REGRESSIONTEXT "
            '    str2 = "FROM ASSAYANALYTES, CONFIGREGRESSIONTYPES "
            '    str2 = str2 & "WHERE ASSAYANALYTES.REGRESSIONIDENTIFIER = CONFIGREGRESSIONTYPES.REGRESSIONID "
            '    str3 = "AND (((ASSAYANALYTES.STUDYID)=" & wStudyID & "));"
            'End If
            'strSQL = str1 & str2 & str3
            '''''''''''''''''''Console.WriteLine(strSQL)

            'rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'rs.ActiveConnection = Nothing

            int1 = FindRow("Regression", tblWatsonAnalRefTable, "Item")
            int2 = FindRow("Weighting", tblWatsonAnalRefTable, "Item")
            int3 = 0

            '****

            'Dim dvT As System.Data.DataView = New DataView(tblRegCon)
            'Dim dvT As System.Data.DataView = New DataView(tblRegConAll)
            '20151220 LEE: No. Record only accepted regressions, not all regressions
            Dim dvT As System.Data.DataView = New DataView(tblRegCon)
            For Count1 = 1 To ctAnalytes

                'If boolIncludePSAE Then
                '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0"
                'Else
                '    str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID <> 3"
                'End If
                str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1) & " AND RUNTYPEID > 0 AND ANALYTEID = " & arrAnalytes(2, Count1)

                '20160927 LEE: This filter needs to include matrix (SAMPLETYPEID), also eventually Group
                '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
                '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX, 17=intOrder, 18=CALIBRSET
                str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND RUNTYPEID > 0 AND ANALYTEID = " & arrAnalytes(2, Count1) & " AND SAMPLETYPEID = '" & arrAnalytes(16, Count1) & "'"
                'str1 = "STUDYID = " & wStudyID & " AND ANALYTEINDEX = " & arrAnalytes(3, Count1) & " AND MASTERASSAYID = " & arrAnalytes(12, Count1)
                dvT.RowFilter = str1

                ''find number of table rows
                '****

                Dim intNumRegr As Short
                Dim tblRunID As System.Data.DataTable
                Dim intTRows As Short
                Dim intRP As Short
                Dim strW As String
                Dim strReg As String

                intTRows = 0
                tblRunID = dvT.ToTable("b", True, "RUNID")
                intTRows = tblRunID.Rows.Count
                '****

                Dim tblT As System.Data.DataTable = dvT.ToTable("a", True, "REGRESSIONPARAMETERID")
                intRP = tblT.Rows.Count

                'determine if there is more than one regr type
                Dim tblSR As System.Data.DataTable = dvT.ToTable("sr", True, "REGRESSIONTEXT", "WEIGHTINGFACTOR")
                intNumRegr = tblSR.Rows.Count

                For Count2 = 1 To intNumRegr
                    'report regression
                    'var1 = rs.Fields("REGRESSIONTEXT").Value
                    var1 = tblSR.Rows(Count2 - 1).Item("REGRESSIONTEXT")
                    'report weighting
                    'var2 = rs.Fields("WEIGHTINGFACTOR").Value
                    var2 = tblSR.Rows(Count2 - 1).Item("WEIGHTINGFACTOR")
                    str1 = GetWt(NZ(var2, 1))
                    If Count2 = 1 Then
                        strReg = var1
                        strW = str1
                    Else
                        strReg = strReg & ", " & var1
                        strW = strW & ", " & str1
                    End If
                Next

                '****

                tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int1).Item(Count1) = strReg
                tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()
                tblWatsonAnalRefTable.Rows.Item(int2).BeginEdit()
                tblWatsonAnalRefTable.Rows.Item(int2).Item(Count1) = strW
                tblWatsonAnalRefTable.Rows.Item(int2).EndEdit()
            Next

            'For Count1 = 1 To ctAnalytes
            '    'str1 = "ANALYTEINDEX=" & arrAnalytes(3, Count1)
            '    str1 = "ANALYTEID=" & arrAnalytes(2, Count1)

            '    int3 = int3 + 1
            '    rs.Filter = str1
            '    intRows = rs.RecordCount
            '    int4 = 1
            '    Do Until rs.EOF
            '        'report regression
            '        var1 = rs.Fields("REGRESSIONTEXT").Value
            '        'report weighting
            '        var2 = rs.Fields("WEIGHTINGFACTOR").Value
            '        str1 = GetWt(NZ(var2, 1))
            '        If int4 = 1 Then
            '            strReg = var1
            '            strW = str1
            '        Else
            '            strReg = strReg & ", " & var1
            '            strW = strW & ", " & str1
            '        End If

            '        rs.MoveNext()
            '        int4 = int4 + 1
            '    Loop
            '    tblWatsonAnalRefTable.Rows.Item(int1).BeginEdit()
            '    tblWatsonAnalRefTable.Rows.Item(int1).Item(int3) = strReg
            '    tblWatsonAnalRefTable.Rows.Item(int1).EndEdit()
            '    tblWatsonAnalRefTable.Rows.Item(int2).BeginEdit()
            '    tblWatsonAnalRefTable.Rows.Item(int2).Item(int3) = strW
            '    tblWatsonAnalRefTable.Rows.Item(int2).EndEdit()

            '    rs.Filter = ""
            'Next
            'rs.Close()
            ''***End here 5

            If rsF.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsF.Close()
            End If
            If rs4.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs4.Close()
            End If
            If rs1.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs1.Close()
            End If
            If rs20.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs20.Close()
            End If
            If rs3.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs3.Close()
            End If


            rsF = Nothing
            '***End here 7

            'prepare heading in TableConfig
            'do something here???

            str1 = "Retrieving Watson Data...22 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            '****begin preparing sample tables


            '*** End tblSampleDesign


            str1 = "Retrieving Watson Data...23 " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            ctPB = ctPB + 1
            If ctPB > frmH.pb1.Maximum Then
                ctPB = 1
            End If
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            Dim rsReassay As New ADODB.Recordset
            'open reassay recordset
            'NOTE DecisionCode should be allowed to be anything
            If boolANSI Then
                str1 = "SELECT SAMPRESCONFLICTDEC.ANALYTEID, SAMPRESCONFLICTDEC.DESIGNSAMPLEID, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTCHOICES.RUNID, SAMPRESCONFLICTCHOICES.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTS.STUDYID "
                str2 = "FROM (SAMPRESCONFLICTDEC INNER JOIN SAMPLERESULTS ON (SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTS.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTS.ANALYTEID) AND (SAMPRESCONFLICTDEC.STUDYID = SAMPLERESULTS.STUDYID)) INNER JOIN SAMPRESCONFLICTCHOICES ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTS.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTS.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.DECISIONCODE)>0) AND ((SAMPLERESULTS.STUDYID)=" & wStudyID & "));"
            Else
                str1 = "SELECT SAMPRESCONFLICTDEC.ANALYTEID, SAMPRESCONFLICTDEC.DESIGNSAMPLEID, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTCHOICES.RUNID, SAMPRESCONFLICTCHOICES.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTS.STUDYID "
                str2 = "FROM SAMPRESCONFLICTDEC, SAMPLERESULTS, SAMPRESCONFLICTCHOICES "
                str2 = str2 & "WHERE ((SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTS.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTS.ANALYTEID) AND (SAMPRESCONFLICTDEC.STUDYID = SAMPLERESULTS.STUDYID)) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTS.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTS.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) "
                str3 = "AND (((SAMPRESCONFLICTDEC.DECISIONCODE)>0) AND ((SAMPLERESULTS.STUDYID)=" & wStudyID & "));"
            End If

            If boolAccess Then
                str1 = "SELECT SAMPRESCONFLICTDEC.ANALYTEID, SAMPRESCONFLICTDEC.DESIGNSAMPLEID, SAMPRESCONFLICTDEC.DECISIONCODE, SAMPRESCONFLICTCHOICES.RUNID, SAMPRESCONFLICTCHOICES.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTS.STUDYID "
                str2 = "FROM (SAMPRESCONFLICTDEC INNER JOIN SAMPLERESULTS ON (SAMPRESCONFLICTDEC.DESIGNSAMPLEID = SAMPLERESULTS.DESIGNSAMPLEID) AND (SAMPRESCONFLICTDEC.ANALYTEID = SAMPLERESULTS.ANALYTEID) AND (SAMPRESCONFLICTDEC.STUDYID = SAMPLERESULTS.STUDYID)) INNER JOIN SAMPRESCONFLICTCHOICES ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTS.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTS.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.DECISIONCODE)>0) AND ((SAMPLERESULTS.STUDYID)=" & wStudyID & "));"
            Else
                str1 = "SELECT " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID, " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".SAMPRESCONFLICTCHOICES.RUNID, " & strSchema & ".SAMPRESCONFLICTCHOICES.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTS.STUDYID "
                str2 = "FROM (" & strSchema & ".SAMPRESCONFLICTDEC INNER JOIN " & strSchema & ".SAMPLERESULTS ON (" & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID = " & strSchema & ".SAMPLERESULTS.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID = " & strSchema & ".SAMPLERESULTS.STUDYID)) INNER JOIN " & strSchema & ".SAMPRESCONFLICTCHOICES ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE)>0) AND ((" & strSchema & ".SAMPLERESULTS.STUDYID)=" & wStudyID & "));"
            End If

            strSQL = str1 & str2 & str3

            ''''Console.WriteLine("rsReassay: " & strSQL)
            rsReassay.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            tblReassay.Clear()
            tblReassay.AcceptChanges()
            tblReassay.BeginLoadData()
            daDoPr.Fill(tblReassay, rsReassay)
            tblReassay.EndLoadData()

            str1 = "Retrieving Watson Data...23a " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            str1 = str1 & ChrW(10) & ChrW(10) & "...If the study is large, this step may take a few moments..."
            frmH.lblProgress.Text = str1
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()

            Cursor.Current = Cursors.WaitCursor

            System.Windows.Forms.Application.DoEvents()


            '*** Create FindNomcConc query

            If boolAccess Then
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.STUDYID "
                str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ASSAYANALYTES ON (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

                '20160816 LEE
                'this query was not returning custom sample types
                'ASSAYANALYTEKNOWN.KNOWNTYPE: limited sample types to QC, STANDARD, STABILITY
                'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND: KNOWNTYPE + UNKNOWN + custom sample types
                'Solution: remove join between KNOWNTYPE and RUNSAMPLEKIND
                'add additional parameters to FindNomConc
                str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.STUDYID "
                str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ASSAYANALYTES ON (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID "
                str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"


            Else
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
                str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND " & strSchema & ".(ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN " & strSchema & ".GLOBALANALYTES ON " & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

                '20160816 LEE
                'this query was not returning custom sample types
                'ASSAYANALYTEKNOWN.KNOWNTYPE: limited sample types to QC, STANDARD, STABILITY
                'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND: KNOWNTYPE + UNKNOWN + custom sample types
                'Solution: remove join between KNOWNTYPE and RUNSAMPLEKIND
                'add additional parameters to FindNomConc
                str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
                str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN " & strSchema & ".GLOBALANALYTES ON " & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID "
                str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

            End If

            '20161227 LEE: Remember this also gets called in AssignedSamples when ChangeStudy

            strSQL = str1 & str2 & str3 & str4

            ''Console.WriteLine("tblConcLevelsForAssayIDs: " & strSQL)

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            tblConcLevelsForAssayIDs.Clear()
            tblConcLevelsForAssayIDs.AcceptChanges()
            tblConcLevelsForAssayIDs.BeginLoadData()
            daDoPr.Fill(tblConcLevelsForAssayIDs, rs)
            tblConcLevelsForAssayIDs.EndLoadData()

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If

            '**** End FindNomConc query




            '** Load all Applicable Rows for Table
            'NDL 12-Feb-2016 We have the following requirements to this Query:
            '(a) Rows have to be in the Table SAMPLERESULTSCONFLICT
            '(b) The TimeStamp in for the Decision must be the same in SAMPRESCONFLICTDEC, SAMPLECONFLICTCHOICES, and SAMPLERESULTS
            If boolAccess Then
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.DESIGNSAMPLEID, DESIGNSAMPLE.USERSAMPLEID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.STUDYID, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPRESCONFLICTDEC.DECISIONCODE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.ALIQUOTFACTOR, SAMPLERESULTS.RUNID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.VISITTEXT "
                str2 = "FROM ((((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID)) ON (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID)) INNER JOIN (DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) ON (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN (((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAY ON (ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANALYTICALRUNSAMPLE ON (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID) AND (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID)) ON (ANARUNANALYTERESULTS.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUN.STUDYID) AND (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER)) ON (SAMPLERESULTSCONFLICT.RUNID = ANALYTICALRUN.RUNID) AND (SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = ASSAYANALYTES.ANALYTEID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECTGROUP.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) "
                str3 = "WHERE ((SAMPRESCONFLICTDEC.REASSAYREASON <> 'NA') AND (ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS = 3) AND ((SAMPRESCONFLICTDEC.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160216 LEE: Access won't allow the previous query to be viewed in design view.
                'This is query previous to adding DESIGNSUBJECTTREATMENT.WEEK
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.DESIGNSAMPLEID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.STUDYID, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPRESCONFLICTDEC.DECISIONCODE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.ALIQUOTFACTOR, SAMPLERESULTS.RUNID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP "
                str2 = "FROM ((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTDEC.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN (((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANALYTICALRUNSAMPLE ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUN.STUDYID) AND (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNANALYTERESULTS.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON (ASSAYANALYTES.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (SAMPLERESULTSCONFLICT.RUNID = ANALYTICALRUN.RUNID)"
                str3 = "WHERE ((SAMPRESCONFLICTDEC.REASSAYREASON <> 'NA') AND (ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS = 3) AND ((SAMPRESCONFLICTDEC.STUDYID) = " & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER "

                '20160216 LEE: Add  DESIGNSUBJECTTREATMENT.WEEK back in
                'This query can be viewed in Access design view.
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.DESIGNSAMPLEID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.STUDYID, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPRESCONFLICTDEC.DECISIONCODE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.ALIQUOTFACTOR, SAMPLERESULTS.RUNID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECTTREATMENT.WEEK "
                str2 = "FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN (((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANALYTICALRUNSAMPLE ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUN.STUDYID) AND (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNANALYTERESULTS.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON (SAMPLERESULTSCONFLICT.ANALYTEID = ASSAYANALYTES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (SAMPLERESULTSCONFLICT.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160217 LEE: Added back DESIGNSAMPLE.USERSAMPLEID and DESIGNSUBJECTTREATMENT.VISITTEXT
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.DESIGNSAMPLEID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.STUDYID, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPRESCONFLICTDEC.DECISIONCODE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.ALIQUOTFACTOR, SAMPLERESULTS.RUNID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.USERSAMPLEID, DESIGNSUBJECTTREATMENT.VISITTEXT "
                str2 = "FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN (((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANALYTICALRUNSAMPLE ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUN.STUDYID) AND (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNANALYTERESULTS.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON (SAMPLERESULTSCONFLICT.ANALYTEID = ASSAYANALYTES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (SAMPLERESULTSCONFLICT.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160219 LEE: Added WEEK to sort
                '20160219 LEE: 
                ' - there were duplicate instances of RUNID in orignal query,  I see you use both in different Unique tables, will leave for now
                ' - there were duplicate instances of ALIQUOTFACTOR in original query, I see you use both in different Unique tables, will leave for now
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.DESIGNSAMPLEID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.STUDYID, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPRESCONFLICTDEC.DECISIONCODE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CONCENTRATION, SAMPLERESULTS.ALIQUOTFACTOR, ANARUNANALYTERESULTS.RUNID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.USERSAMPLEID, DESIGNSUBJECTTREATMENT.VISITTEXT, SAMPLERESULTS.RUNID "
                str2 = "FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN (((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANALYTICALRUNSAMPLE ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUN.STUDYID) AND (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNANALYTERESULTS.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON (SAMPLERESULTSCONFLICT.ANALYTEID = ASSAYANALYTES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (SAMPLERESULTSCONFLICT.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"
                'str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160502 LEE:  added start times
                ', DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND
                '20161108 LEE: Changed to "as SAMPLERESULTS_CONCENTRATION", "as SAMPLERESULTS_ALIQUOTFACTOR", "as SAMPLERESULTS_RUNID", "as ANARUNANALYTERESULTS_RUNID", "AS ANARUNANALYTERESULTS_CONCENTRATION", "as ARS_ALIQUOTFACTOR" to account for Oracle syntax
                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.DESIGNSAMPLEID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.STUDYID, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPRESCONFLICTDEC.DECISIONCODE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNANALYTERESULTS.CONCENTRATION AS AR_CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR as ARS_ALIQUOTFACTOR, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CONCENTRATION as SAMPLERESULTS_CONCENTRATION, SAMPLERESULTS.ALIQUOTFACTOR as SAMPLERESULTS_ALIQUOTFACTOR, ANARUNANALYTERESULTS.RUNID AS ANARUNANALYTERESULTS_RUNID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.USERSAMPLEID, DESIGNSUBJECTTREATMENT.VISITTEXT, SAMPLERESULTS.RUNID as SAMPLERESULTS_RUNID, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN (((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANALYTICALRUNSAMPLE ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUN.STUDYID) AND (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNANALYTERESULTS.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON (SAMPLERESULTSCONFLICT.ANALYTEID = ASSAYANALYTES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (SAMPLERESULTSCONFLICT.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"


                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,

                str1 = "SELECT DISTINCT SAMPLERESULTSCONFLICT.ANALYTEID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSAMPLE.DESIGNSAMPLEID, DESIGNSAMPLE.TREATMENTEVENTID, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, DESIGNSAMPLE.ENDSECOND, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, SAMPLERESULTSCONFLICT.STUDYID, SAMPLERESULTSCONFLICT.ORIGINALVALUE, SAMPRESCONFLICTDEC.REASSAYREASON, SAMPRESCONFLICTDEC.REASSAYCONCREASON, SAMPRESCONFLICTDEC.DECISIONCODE, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNANALYTERESULTS.CONCENTRATION AS AR_CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, ANALYTICALRUNANALYTES.NM, ANALYTICALRUNANALYTES.VEC, Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ARS_ALIQUOTFACTOR, SAMPLERESULTS.CALIBRATIONRANGEFLAG, SAMPLERESULTS.CONCENTRATION as SAMPLERESULTS_CONCENTRATION, Round([SAMPLERESULTS].[ALIQUOTFACTOR]," & intDFDec & ") AS SAMPLERESULTS_ALIQUOTFACTOR, ANARUNANALYTERESULTS.RUNID AS ANARUNANALYTERESULTS_RUNID, SAMPRESCONFLICTDEC.RECORDTIMESTAMP, SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, SAMPLERESULTS.ACCEPTANCETIMESTAMP, SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.USERSAMPLEID, DESIGNSUBJECTTREATMENT.VISITTEXT, SAMPLERESULTS.RUNID as SAMPLERESULTS_RUNID, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM (((SAMPLERESULTS INNER JOIN (SAMPLERESULTSCONFLICT INNER JOIN (SAMPRESCONFLICTCHOICES INNER JOIN SAMPRESCONFLICTDEC ON (SAMPRESCONFLICTCHOICES.STUDYID = SAMPRESCONFLICTDEC.STUDYID) AND (SAMPRESCONFLICTCHOICES.ANALYTEID = SAMPRESCONFLICTDEC.ANALYTEID) AND (SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (SAMPLERESULTSCONFLICT.STUDYID = SAMPRESCONFLICTCHOICES.STUDYID) AND (SAMPLERESULTSCONFLICT.ANALYTEID = SAMPRESCONFLICTCHOICES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (SAMPLERESULTS.ANALYTEID = SAMPLERESULTSCONFLICT.ANALYTEID) AND (SAMPLERESULTS.ACCEPTANCETIMESTAMP = SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (SAMPLERESULTS.STUDYID = SAMPLERESULTSCONFLICT.STUDYID) AND (SAMPLERESULTS.DESIGNSAMPLEID = SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((DESIGNSAMPLE INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (SAMPLERESULTS.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (SAMPLERESULTS.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN (ANARUNANALYTERESULTS INNER JOIN (((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAY.ASSAYID)) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN ANALYTICALRUNSAMPLE ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (ANARUNANALYTERESULTS.STUDYID = ANALYTICALRUN.STUDYID) AND (ANARUNANALYTERESULTS.RUNID = ANALYTICALRUN.RUNID) AND (ANARUNANALYTERESULTS.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX)) ON (SAMPLERESULTSCONFLICT.ANALYTEID = ASSAYANALYTES.ANALYTEID) AND (SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (SAMPLERESULTSCONFLICT.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, SAMPLERESULTSCONFLICT.ANALYTEID, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNANALYTERESULTS.RUNID, SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                'anarunanalyteresults.runid
                'ANALYTICALRUNSAMPLE.ALIQUOTFACTOR


            Else
                '*************** NDL ********** DO this later

                '20160216 LEE: Add  DESIGNSUBJECTTREATMENT.WEEK back in
                'This query can be viewed in Access design view.
                str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK "
                str2 = "FROM (((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN (" & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANALYTICALRUNSAMPLE ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNANALYTERESULTS.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160217 LEE: Added back DESIGNSAMPLE.USERSAMPLEID and DESIGNSUBJECTTREATMENT.VISITTEXT
                str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".SAMPLERESULTS.RUNID, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT "
                str2 = "FROM (((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN (" & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANALYTICALRUNSAMPLE ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNANALYTERESULTS.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160219 LEE: Added WEEK to sort
                str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CONCENTRATION, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".SAMPLERESULTS.RUNID "
                str2 = "FROM (((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN (" & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANALYTICALRUNSAMPLE ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNANALYTERESULTS.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"

                '20160502 LEE:  added start times
                ', DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.STARTSECOND
                '20161108 LEE: Changed to "as SAMPLERESULTS_CONCENTRATION", "as SAMPLERESULTS_ALIQUOTFACTOR", "as SAMPLERESULTS_RUNID", "as ANARUNANALYTERESULTS_RUNID", "AS ANARUNANALYTERESULTS_CONCENTRATION", "as ARS_ALIQUOTFACTOR" to account for Oracle syntax
                str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION AS AR_CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR as ARS_ALIQUOTFACTOR, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CONCENTRATION as SAMPLERESULTS_CONCENTRATION, " & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR as SAMPLERESULTS_ALIQUOTFACTOR, " & strSchema & ".ANARUNANALYTERESULTS.RUNID AS ANARUNANALYTERESULTS_RUNID, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".SAMPLERESULTS.RUNID as SAMPLERESULTS_RUNID, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM (((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN (" & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANALYTICALRUNSAMPLE ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNANALYTERESULTS.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"


                '20171124 LEE:
                'Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR,
                str1 = "SELECT DISTINCT " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".DESIGNSAMPLE.TREATMENTEVENTID, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER, " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID, " & strSchema & ".SAMPLERESULTSCONFLICT.ORIGINALVALUE, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON, " & strSchema & ".SAMPRESCONFLICTDEC.REASSAYCONCREASON, " & strSchema & ".SAMPRESCONFLICTDEC.DECISIONCODE, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION AS AR_CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANALYTICALRUNANALYTES.NM, " & strSchema & ".ANALYTICALRUNANALYTES.VEC, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR as ARS_ALIQUOTFACTOR, " & strSchema & ".SAMPLERESULTS.CALIBRATIONRANGEFLAG, " & strSchema & ".SAMPLERESULTS.CONCENTRATION as SAMPLERESULTS_CONCENTRATION, ROUND(" & strSchema & ".SAMPLERESULTS.ALIQUOTFACTOR," & intDFDec & ") AS SAMPLERESULTS_ALIQUOTFACTOR, " & strSchema & ".ANARUNANALYTERESULTS.RUNID AS ANARUNANALYTERESULTS_RUNID, " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP, " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP, " & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP, " & strSchema & ".SAMPLERESULTSCONFLICT.RECORDTIMESTAMP, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".SAMPLERESULTS.RUNID as SAMPLERESULTS_RUNID, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.STARTSECOND "
                str2 = "FROM (((" & strSchema & ".SAMPLERESULTS INNER JOIN (" & strSchema & ".SAMPLERESULTSCONFLICT INNER JOIN (" & strSchema & ".SAMPRESCONFLICTCHOICES INNER JOIN " & strSchema & ".SAMPRESCONFLICTDEC ON (" & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID = " & strSchema & ".SAMPRESCONFLICTDEC.STUDYID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTDEC.ANALYTEID) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP = " & strSchema & ".SAMPRESCONFLICTDEC.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTDEC.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = " & strSchema & ".SAMPRESCONFLICTCHOICES.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID = " & strSchema & ".SAMPRESCONFLICTCHOICES.DESIGNSAMPLEID)) ON (" & strSchema & ".SAMPLERESULTS.ANALYTEID = " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTS.ACCEPTANCETIMESTAMP = " & strSchema & ".SAMPRESCONFLICTCHOICES.RECORDTIMESTAMP) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID) AND (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".SAMPLERESULTSCONFLICT.DESIGNSAMPLEID)) INNER JOIN ((" & strSchema & ".DESIGNSAMPLE INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) ON (" & strSchema & ".SAMPLERESULTS.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".SAMPLERESULTS.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN (" & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) INNER JOIN " & strSchema & ".ANALYTICALRUNSAMPLE ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANARUNANALYTERESULTS.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX)) ON (" & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID = " & strSchema & ".ASSAYANALYTES.ANALYTEID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.STUDYID = ANALYTICALRUN.STUDYID) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".SAMPLERESULTSCONFLICT.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) "
                str3 = "WHERE (((" & strSchema & ".SAMPRESCONFLICTDEC.REASSAYREASON)<>'NA') AND ((" & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS)=3) AND ((" & strSchema & ".SAMPRESCONFLICTDEC.STUDYID)=" & wStudyID & ")) "
                str4 = "ORDER BY " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".SAMPLERESULTSCONFLICT.ANALYTEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ANARUNANALYTERESULTS.RUNID, " & strSchema & ".SAMPLERESULTSCONFLICT.RUNSAMPLESEQUENCENUMBER;"


                'ANARUNANALYTERESULTS.RUNID
                '.RUNID()
                'sampleresults.concentration


                'USERSAMPLEID

            End If

            strSQL = str1 & str2 & str3 & str4
            '20171106 LEE: Microsoft Access memory error here for Alturas IT001101
            'Console.WriteLine("tblRepeatAllRunSamples: " & strSQL)
            Dim rsRepeatAllRunSamples As New ADODB.Recordset
            rsRepeatAllRunSamples.Filter = ""
            If rsRepeatAllRunSamples.State = ADODB.ObjectStateEnum.adStateOpen Then
                rsRepeatAllRunSamples.Close()
            End If

            rsRepeatAllRunSamples.CursorLocation = CursorLocationEnum.adUseClient

            Try
                rsRepeatAllRunSamples.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                var1 = var1
            End Try

            rsRepeatAllRunSamples.ActiveConnection = Nothing

            'Create Table of All Samples
            Try
                tblRepeatAllRunSamples.Clear()
                tblRepeatAllRunSamples.AcceptChanges()
                tblRepeatAllRunSamples.BeginLoadData()
                daDoPr.Fill(tblRepeatAllRunSamples, rsRepeatAllRunSamples)
                tblRepeatAllRunSamples.EndLoadData()
                rsRepeatAllRunSamples.Close()
                rsRepeatAllRunSamples = Nothing
            Catch ex As Exception
                str1 = "A query to retrieve samples related to the Repeat Samples table was too large to execute." & ChrW(10) & ChrW(10)
                str1 = str1 & "The Repeat Samples table will not be able to be generated in this study." & ChrW(10) & ChrW(10)
                str1 = str1 & ex.Message
                MsgBox(str1, vbInformation, "Problem...")
                var1 = var1
            End Try


            '20180227 LEE:
            'LI-00016 Study POP01 Furosemide Human Plasma has several original and repeat samples marked incorrectly in SAMPLERESULTSCONFLICT.ORIGINALVALUE (Y and N are switched).
            'This is obvious because the Reassay RunID (e.g. 3) is before the Original RunID (e.g. 7)
            'There must be some problem with the query because actaul Reassay Watson table provided by LI-00016 shows the correct assignment.
            'need to check here to ensure correct assignment
            Try
                Call AdjustReassayOriginalValue()
            Catch ex As Exception
                var1 = var1 'debug
            End Try



            'Make ISR table
            '20160717 LEE: Nick may want to evaluate this query. I had to get analyteid from assayanalytes
            If boolCanDoISR Then

                Try

                    If boolAccess Then
                        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.ANALYSISTYPE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, CONFIGGENDER.GENDER, DESIGNSAMPLE.TIMETEXT, DESIGNSUBJECTTREATMENT.VISITTEXT, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.WEEK, ANALYTICALRUNSAMPLE.RUNID, DESIGNSUBJECT.GENDERID, DESIGNSUBJECTGROUP.SUBJECTGROUPID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTSECOND "
                        str2 = "FROM (((((((((ANALYTICALRUNSAMPLE INNER JOIN DESIGNSAMPLE ON (ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (ANALYTICALRUNSAMPLE.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID)) INNER JOIN DESIGNTREATMENT ON (DESIGNSUBJECTTREATMENT.TREATMENTID = DESIGNTREATMENT.TREATMENTID) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNTREATMENT.STUDYID)) INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) LEFT JOIN CONFIGGENDER ON DESIGNSUBJECT.GENDERID = CONFIGGENDER.GENDERID) INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUN.STUDYID = ASSAYANALYTES.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & "));"

                        '20170924 LEE: Added , ASSAYANALYTES.NM, ASSAYANALYTES.VEC to correctly report BQL/AQL
                        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.ANALYSISTYPE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, CONFIGGENDER.GENDER, DESIGNSAMPLE.TIMETEXT, DESIGNSUBJECTTREATMENT.VISITTEXT, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, DESIGNSUBJECTTREATMENT.WEEK, DESIGNSUBJECTTREATMENT.WEEK, ANALYTICALRUNSAMPLE.RUNID, DESIGNSUBJECT.GENDERID, DESIGNSUBJECTGROUP.SUBJECTGROUPID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTSECOND, ASSAYANALYTES.NM, ASSAYANALYTES.VEC "
                        str2 = "FROM (((((((((ANALYTICALRUNSAMPLE INNER JOIN DESIGNSAMPLE ON (ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (ANALYTICALRUNSAMPLE.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID)) INNER JOIN DESIGNTREATMENT ON (DESIGNSUBJECTTREATMENT.TREATMENTID = DESIGNTREATMENT.TREATMENTID) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNTREATMENT.STUDYID)) INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) LEFT JOIN CONFIGGENDER ON DESIGNSUBJECT.GENDERID = CONFIGGENDER.GENDERID) INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUN.STUDYID = ASSAYANALYTES.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & "));"

                        '20180227 LEE
                        'need to add , DESIGNSAMPLE.USERSAMPLEID
                        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.ANALYSISTYPE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, CONFIGGENDER.GENDER, DESIGNSAMPLE.TIMETEXT, DESIGNSUBJECTTREATMENT.VISITTEXT, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, DESIGNSUBJECTTREATMENT.WEEK, ANALYTICALRUNSAMPLE.RUNID, DESIGNSUBJECT.GENDERID, DESIGNSUBJECTGROUP.SUBJECTGROUPID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTSECOND, ASSAYANALYTES.NM, ASSAYANALYTES.VEC, DESIGNSAMPLE.USERSAMPLEID "
                        str2 = "FROM (((((((((ANALYTICALRUNSAMPLE INNER JOIN DESIGNSAMPLE ON (ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) AND (ANALYTICALRUNSAMPLE.STUDYID = DESIGNSAMPLE.STUDYID)) INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID)) INNER JOIN DESIGNTREATMENT ON (DESIGNSUBJECTTREATMENT.TREATMENTID = DESIGNTREATMENT.TREATMENTID) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNTREATMENT.STUDYID)) INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) LEFT JOIN CONFIGGENDER ON DESIGNSUBJECT.GENDERID = CONFIGGENDER.GENDERID) INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUN.STUDYID = ASSAYANALYTES.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & "));"
                        'USERSAMPLEID

                        '20190117 LEE
                        'query was returning replicate records. Needed to tweek the query
                        'query was also returning deactivated (eliminated) samples. This has been fixed
                        str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.ANALYSISTYPE, CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.ANALYTEID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANARUNANALYTERESULTS.CONCENTRATION, ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, DESIGNSUBJECT.DESIGNSUBJECTID, DESIGNSUBJECT.DESIGNSUBJECTTAG, DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, DESIGNTREATMENT.TREATMENTID, DESIGNTREATMENT.TREATMENTDESC, DESIGNSAMPLE.STARTDAY, DESIGNSAMPLE.STARTHOUR, DESIGNSAMPLE.STARTMINUTE, DESIGNSAMPLE.ENDDAY, DESIGNSAMPLE.ENDHOUR, DESIGNSAMPLE.ENDMINUTE, CONFIGGENDER.GENDER, DESIGNSAMPLE.TIMETEXT, DESIGNSUBJECTTREATMENT.VISITTEXT, ANALYTICALRUNSAMPLE.STUDYID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, DESIGNSUBJECTTREATMENT.Week, ANALYTICALRUNSAMPLE.RUNID, DESIGNSUBJECT.GENDERID, DESIGNSUBJECTGROUP.SUBJECTGROUPID, DESIGNSAMPLE.ENDSECOND, DESIGNSAMPLE.STARTSECOND, ASSAYANALYTES.NM, ASSAYANALYTES.VEC, DESIGNSAMPLE.USERSAMPLEID, ANARUNANALYTERESULTS.ELIMINATEDFLAG "
                        str2 = "FROM (((((((((ANALYTICALRUNSAMPLE INNER JOIN DESIGNSAMPLE ON (ANALYTICALRUNSAMPLE.STUDYID = DESIGNSAMPLE.STUDYID) AND (ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN DESIGNSUBJECT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTID = DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN DESIGNSUBJECTGROUP ON (DESIGNSUBJECT.STUDYID = DESIGNSUBJECTGROUP.STUDYID) AND (DESIGNSUBJECT.SUBJECTGROUPID = DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN DESIGNSUBJECTTREATMENT ON (DESIGNSAMPLE.STUDYID = DESIGNSUBJECTTREATMENT.STUDYID) AND (DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN DESIGNTREATMENT ON (DESIGNSUBJECTTREATMENT.STUDYID = DESIGNTREATMENT.STUDYID) AND (DESIGNSUBJECTTREATMENT.TREATMENTKEY = DESIGNTREATMENT.TREATMENTKEY) AND (DESIGNSUBJECTTREATMENT.TREATMENTID = DESIGNTREATMENT.TREATMENTID)) INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) LEFT JOIN CONFIGGENDER ON DESIGNSUBJECT.GENDERID = CONFIGGENDER.GENDERID) INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAYANALYTES ON (ASSAYANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUN.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID)) INNER JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((ANARUNANALYTERESULTS.ELIMINATEDFLAG)<>'Y'));"

                    Else
                        str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".CONFIGGENDER.GENDER, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".DESIGNSAMPLE.STARTSECOND "
                        str2 = "FROM (((((((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNTREATMENT ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTID = " & strSchema & ".DESIGNTREATMENT.TREATMENTID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) LEFT JOIN " & strSchema & ".CONFIGGENDER ON " & strSchema & ".DESIGNSUBJECT.GENDERID = " & strSchema & ".CONFIGGENDER.GENDERID) INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ". DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & "));"

                        '20170924 LEE: Added , ASSAYANALYTES.NM, ASSAYANALYTES.VEC to correctly report BQL/AQL
                        str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".CONFIGGENDER.GENDER, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC "
                        str2 = "FROM (((((((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNTREATMENT ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTID = " & strSchema & ".DESIGNTREATMENT.TREATMENTID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) LEFT JOIN " & strSchema & ".CONFIGGENDER ON " & strSchema & ".DESIGNSUBJECT.GENDERID = " & strSchema & ".CONFIGGENDER.GENDERID) INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ". DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & "));"

                        '20180227 LEE
                        'need to add , DESIGNSAMPLE.USERSAMPLEID
                        str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".CONFIGGENDER.GENDER, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID "
                        str2 = "FROM (((((((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNTREATMENT ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTID = " & strSchema & ".DESIGNTREATMENT.TREATMENTID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) LEFT JOIN " & strSchema & ".CONFIGGENDER ON " & strSchema & ".DESIGNSUBJECT.GENDERID = " & strSchema & ".CONFIGGENDER.GENDERID) INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ". DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & "));"

                        ''20190117 LEE
                        ''query was returning replicate records. Needed to tweek the query
                        ''query was also returning deactivated (eliminated) samples. This has been fixed
                        str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".CONFIGGENDER.GENDER, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".DESIGNSUBJECTTREATMENT.Week, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".DESIGNSAMPLE.STARTSECOND, " & strSchema & ".ASSAYANALYTES.NM, " & strSchema & ".ASSAYANALYTES.VEC, " & strSchema & ".DESIGNSAMPLE.USERSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG "
                        str2 = "FROM (((((((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID)) INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID) AND (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID) AND (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY)) INNER JOIN " & strSchema & ".DESIGNTREATMENT ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNTREATMENT.STUDYID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTID = " & strSchema & ".DESIGNTREATMENT.TREATMENTID)) INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) LEFT JOIN " & strSchema & ".CONFIGGENDER ON " & strSchema & ".DESIGNSUBJECT.GENDERID = " & strSchema & ".CONFIGGENDER.GENDERID) INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                        str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & ") AND ((" & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG)<>'Y'));"


                    End If

                    'strSchema = "GWatson"

                    'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATIONSTATUS, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID, " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTTAG, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPNAME, " & strSchema & ".DESIGNTREATMENT.TREATMENTID, " & strSchema & ".DESIGNTREATMENT.TREATMENTDESC, " & strSchema & ".DESIGNSAMPLE.STARTDAY, " & strSchema & ".DESIGNSAMPLE.STARTHOUR, " & strSchema & ".DESIGNSAMPLE.STARTMINUTE, " & strSchema & ".DESIGNSAMPLE.ENDDAY, " & strSchema & ".DESIGNSAMPLE.ENDHOUR, " & strSchema & ".DESIGNSAMPLE.ENDMINUTE, " & strSchema & ".CONFIGGENDER.GENDER, " & strSchema & ".DESIGNSAMPLE.TIMETEXT, " & strSchema & ".DESIGNSUBJECTTREATMENT.VISITTEXT, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".DESIGNSUBJECTTREATMENT.WEEK, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".DESIGNSUBJECT.GENDERID, " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID, " & strSchema & ".DESIGNSAMPLE.ENDSECOND, " & strSchema & ".DESIGNSAMPLE.STARTSECOND "
                    'str2 = "FROM ((((((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".DESIGNSAMPLE ON (" & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".DESIGNSAMPLE.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTID = " & strSchema & ".DESIGNSUBJECT.DESIGNSUBJECTID) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTGROUP ON (" & strSchema & ".DESIGNSUBJECT.SUBJECTGROUPID = " & strSchema & ".DESIGNSUBJECTGROUP.SUBJECTGROUPID) AND (" & strSchema & ".DESIGNSUBJECT.STUDYID = " & strSchema & ".DESIGNSUBJECTGROUP.STUDYID)) INNER JOIN " & strSchema & ".DESIGNSUBJECTTREATMENT ON (" & strSchema & ".DESIGNSAMPLE.DESIGNSUBJECTTREATMENTKEY = " & strSchema & ".DESIGNSUBJECTTREATMENT.DESIGNSUBJECTTREATMENTKEY) AND (" & strSchema & ".DESIGNSAMPLE.STUDYID = " & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".DESIGNTREATMENT ON (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTID = " & strSchema & ".DESIGNTREATMENT.TREATMENTID) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.TREATMENTKEY = " & strSchema & ".DESIGNTREATMENT.TREATMENTKEY) AND (" & strSchema & ".DESIGNSUBJECTTREATMENT.STUDYID = " & strSchema & ".DESIGNTREATMENT.STUDYID)) INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) LEFT JOIN " & strSchema & ".CONFIGGENDER ON " & strSchema & ".DESIGNSUBJECT.GENDERID = " & strSchema & ".CONFIGGENDER.GENDERID) INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".DESIGNSAMPLE.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY "
                    'str3 = "WHERE (((" & strSchema & ".ANALYTICALRUNSAMPLE.ANALYSISTYPE)='ISR') AND ((" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)=" & wStudyID & "));"

                    strSQL = str1 & str2 & str3

                    'Console.WriteLine("tblISR: " & strSQL)

                    Dim rsISR As New ADODB.Recordset
                    If rsISR.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rsISR.Close()
                    End If

                    rsISR.CursorLocation = CursorLocationEnum.adUseClient
                    rsISR.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    rsISR.ActiveConnection = Nothing

                    tblISR.Clear()
                    tblISR.AcceptChanges()
                    tblISR.BeginLoadData()
                    daDoPr.Fill(tblISR, rsISR)
                    tblISR.EndLoadData()

                    If tblISR.Columns.Contains("AnalyteDescription") Then
                    Else
                        Dim col10 As New DataColumn
                        col10.ColumnName = "AnalyteDescription"
                        col10.DataType = System.Type.GetType("System.String")
                        tblISR.Columns.Add(col10)
                    End If

                    If tblISR.Columns.Contains("SERIALENDTIME") Then
                    Else
                        Dim col11 As New DataColumn
                        col11.ColumnName = "SERIALENDTIME"
                        col11.DataType = System.Type.GetType("System.Int64")
                        tblISR.Columns.Add(col11)
                    End If

                    If tblISR.Columns.Contains("SERIALSTARTTIME") Then
                    Else
                        Dim col12 As New DataColumn
                        col12.ColumnName = "SERIALSTARTTIME"
                        col12.DataType = System.Type.GetType("System.Int64")
                        tblISR.Columns.Add(col12)
                    End If

                    Count2 = -1

                    Do Until rsISR.EOF
                        'Dim drow As New DataRow
                        Count2 = Count2 + 1
                        'drow = tblSampleDesign.NewRow
                        'drow("Analyte") = "Analyte"
                        tblISR.Rows.Item(Count2).BeginEdit()
                        tblISR.Rows.Item(Count2).Item("AnalyteDescription") = "Analyte"

                        'now calculate SERIALENDTIME
                        vS = NZ(rsISR.Fields("ENDSECOND").Value, 0)
                        vM = NZ(rsISR.Fields("ENDMINUTE").Value, 0)
                        vH = NZ(rsISR.Fields("ENDHOUR").Value, 0)
                        vD = NZ(rsISR.Fields("ENDDAY").Value, 0)

                        'convert to seconds
                        vS = vS * 1
                        vM = vM * 60
                        vH = vH * 60 * 60
                        vD = vD * 60 * 60 * 24
                        sDate = vS + vM + vH + vD

                        tblISR.Rows(Count2).Item("SERIALENDTIME") = sDate

                        'now calculate SERIALSTARTTIME
                        vS = NZ(rsISR.Fields("STARTSECOND").Value, 0)
                        vM = NZ(rsISR.Fields("STARTMINUTE").Value, 0)
                        vH = NZ(rsISR.Fields("STARTHOUR").Value, 0)
                        vD = NZ(rsISR.Fields("STARTDAY").Value, 0)

                        'convert to seconds
                        vS = vS * 1
                        vM = vM * 60
                        vH = vH * 60 * 60
                        vD = vD * 60 * 60 * 24
                        sDate = vS + vM + vH + vD

                        tblISR.Rows(Count2).Item("SERIALSTARTTIME") = sDate
                        tblISR.Rows.Item(Count2).EndEdit()

                        rsISR.MoveNext()

                    Loop

                    rsISR.Close()
                    rsISR = Nothing

                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try



            End If




            '****


            str1 = "Retrieving Watson Data...23b " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If

            Try
                Call AddAnalyteColReportTableAnalytes()

            Catch ex As Exception
                MsgBox("There was a problem executing AddAnalyteColReportTableAnalytes." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            End Try

            str1 = "Retrieving Watson Data...23c " & ctPB
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'Fill tblSampleReceiptWatson

            If boolAccess Then
                str1 = "SELECT DISTINCT Format(SHIPMENT.DATERECEIVED,""mm/dd/yyyy"") AS DATEREC, SHIPMENT.DATERECEIVED, Count(CONTAINERSAMPLE.DESIGNSAMPLEID) AS SAMPLECOUNT, STORAGELOCATION.TEMPERATURE, First(SHIPMENT.COMMENTMEMO) AS FirstOfCOMMENTMEMO "
                str2 = "FROM ((((SHIPMENT INNER JOIN SHIPMENTBOX ON SHIPMENT.SHIPMENTID = SHIPMENTBOX.SHIPMENTID) INNER JOIN LOCATIONCONTAINER ON SHIPMENTBOX.SHIPBOXID = LOCATIONCONTAINER.SHIPBOXID) INNER JOIN [CONTAINER] ON LOCATIONCONTAINER.CONTAINERID = CONTAINER.CONTAINERID) INNER JOIN CONTAINERSAMPLE ON CONTAINER.CONTAINERID = CONTAINERSAMPLE.CONTAINERID) INNER JOIN STORAGELOCATION ON LOCATIONCONTAINER.STORAGELOCID = STORAGELOCATION.STORAGELOCID "
                str3 = "GROUP BY CONTAINERSAMPLE.STUDYID, SHIPMENT.DATERECEIVED, STORAGELOCATION.TEMPERATURE "
                str3 = str3 & "HAVING(((CONTAINERSAMPLE.STUDYID) = " & wStudyID & ") And ((SHIPMENT.DATERECEIVED) Is Not Null)) "
                str4 = "ORDER BY SHIPMENT.DATERECEIVED, Count(CONTAINERSAMPLE.DESIGNSAMPLEID);"
            Else
                'str1 = "SELECT DISTINCT Format(" & strSchema & ".SHIPMENT.DATERECEIVED,""mm/dd/yyyy"") AS DATEREC, " & strSchema & ".SHIPMENT.DATERECEIVED, Count(" & strSchema & ".CONTAINERSAMPLE.DESIGNSAMPLEID) AS SAMPLECOUNT, " & strSchema & ".STORAGELOCATION.TEMPERATURE, First(" & strSchema & ".SHIPMENT.COMMENTMEMO) AS FirstOfCOMMENTMEMO "
                'str2 = "FROM ((((" & strSchema & ".SHIPMENT INNER JOIN " & strSchema & ".SHIPMENTBOX ON " & strSchema & ".SHIPMENT.SHIPMENTID = " & strSchema & ".SHIPMENTBOX.SHIPMENTID) INNER JOIN " & strSchema & ".LOCATIONCONTAINER ON " & strSchema & ".SHIPMENTBOX.SHIPBOXID = " & strSchema & ".LOCATIONCONTAINER.SHIPBOXID) INNER JOIN " & strSchema & ".[CONTAINER] ON " & strSchema & ".LOCATIONCONTAINER.CONTAINERID = " & strSchema & ".CONTAINER.CONTAINERID) INNER JOIN " & strSchema & ".CONTAINERSAMPLE ON " & strSchema & ".CONTAINER.CONTAINERID = " & strSchema & ".CONTAINERSAMPLE.CONTAINERID) INNER JOIN " & strSchema & ".STORAGELOCATION ON " & strSchema & ".LOCATIONCONTAINER.STORAGELOCID = " & strSchema & ".STORAGELOCATION.STORAGELOCID "
                'str3 = "GROUP BY " & strSchema & ".CONTAINERSAMPLE.STUDYID, " & strSchema & ".SHIPMENT.DATERECEIVED, " & strSchema & ".STORAGELOCATION.TEMPERATURE "
                'str3 = str3 & "HAVING(((" & strSchema & ".CONTAINERSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".SHIPMENT.DATERECEIVED) Is Not Null)) "
                'str4 = "ORDER BY " & strSchema & ".SHIPMENT.DATERECEIVED, Count(" & strSchema & ".CONTAINERSAMPLE.DESIGNSAMPLEID);"

                str1 = "SELECT DISTINCT TO_CHAR(" & strSchema & ".SHIPMENT.DATERECEIVED, 'MM-DD-YYYY') AS DATEREC, " & strSchema & ".SHIPMENT.DATERECEIVED, Count(" & strSchema & ".CONTAINERSAMPLE.DESIGNSAMPLEID) AS SAMPLECOUNT, " & strSchema & ".STORAGELOCATION.TEMPERATURE, FIRST_VALUE(" & strSchema & ".SHIPMENT.COMMENTMEMO) OVER(ORDER BY " & strSchema & ".SHIPMENT.COMMENTMEMO ASC) AS FirstOfCOMMENTMEMO "
                str2 = "FROM ((((" & strSchema & ".SHIPMENT INNER JOIN " & strSchema & ".SHIPMENTBOX ON " & strSchema & ".SHIPMENT.SHIPMENTID = " & strSchema & ".SHIPMENTBOX.SHIPMENTID) INNER JOIN " & strSchema & ".LOCATIONCONTAINER ON " & strSchema & ".SHIPMENTBOX.SHIPBOXID = " & strSchema & ".LOCATIONCONTAINER.SHIPBOXID) INNER JOIN " & strSchema & ".CONTAINER ON " & strSchema & ".LOCATIONCONTAINER.CONTAINERID = " & strSchema & ".CONTAINER.CONTAINERID) INNER JOIN " & strSchema & ".CONTAINERSAMPLE ON " & strSchema & ".CONTAINER.CONTAINERID = " & strSchema & ".CONTAINERSAMPLE.CONTAINERID) INNER JOIN " & strSchema & ".STORAGELOCATION ON " & strSchema & ".LOCATIONCONTAINER.STORAGELOCID = " & strSchema & ".STORAGELOCATION.STORAGELOCID "
                str3 = "GROUP BY " & strSchema & ".CONTAINERSAMPLE.STUDYID, " & strSchema & ".SHIPMENT.DATERECEIVED, " & strSchema & ".STORAGELOCATION.TEMPERATURE, " & strSchema & ".SHIPMENT.COMMENTMEMO "
                str3 = str3 & "HAVING(((" & strSchema & ".CONTAINERSAMPLE.STUDYID) = " & wStudyID & ") And ((" & strSchema & ".SHIPMENT.DATERECEIVED) Is Not Null)) "
                str4 = "ORDER BY " & strSchema & ".SHIPMENT.DATERECEIVED;"

            End If

            strSQL = str1 & str2 & str3 & str4

            ''Console.WriteLine("tblSampleReceiptWatson: " & strSQL)

            Dim rsWShipment As New ADODB.Recordset

            Try
                rsWShipment.CursorLocation = CursorLocationEnum.adUseClient
                rsWShipment.Open(strSQL, cn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
                rsWShipment.ActiveConnection = Nothing
                int1 = rsWShipment.RecordCount 'debug

                tblSampleReceiptWatson.Clear()
                tblSampleReceiptWatson.AcceptChanges()
                tblSampleReceiptWatson.BeginLoadData()
                daDoPr.Fill(tblSampleReceiptWatson, rsWShipment)
                tblSampleReceiptWatson.EndLoadData()
                rsWShipment.Close()
                rsWShipment = Nothing

                'total samplecount
                Dim intSRC As Int32 = 0
                Dim intSRCa As Int32 = 0
                For Count1 = 0 To tblSampleReceiptWatson.Rows.Count - 1
                    intSRCa = NZ(tblSampleReceiptWatson.Rows(Count1).Item("SAMPLECOUNT"), 0)
                    intSRC = intSRC + intSRCa
                Next

                '20170127 LEE: DON'T FORMAT NUMBER TO #,###
                'Save command can't convert to integer for some reason
                frmH.txtSRecTotalReportWatson.Text = intSRC

                dv = tblSampleReceiptWatson.DefaultView
                dv.AllowDelete = False
                dv.AllowEdit = False
                dv.AllowNew = False
                frmH.dgvSampleReceiptWatson.DataSource = dv
                frmH.dgvSampleReceiptWatson.Refresh()
                frmH.dgvSampleReceiptWatson.AutoResizeColumns()

                Call InitWatsonSampleReciept()

            Catch ex As Exception
                If rsWShipment.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsWShipment.Close()
                End If
                rsWShipment = Nothing
                'MsgBox("There was a problem opening rsWShipment." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            End Try




            ''Fill tblSampleReceiptWatson
            'If boolAccess Then
            '    str1 = "SELECT SHIPMENTTRANSFER.SHIPMENTID, SHIPMENTTRANSFER.STORAGETEMPERATURE, SHIPMENTTRANSFER.SAMPLECONDITION, SHIPMENTTRANSFER.TRANSFERDATE, SHIPMENTTRANSFER.STUDYID, SHIPMENTTRANSFER.SHIPMENTTRANSFERKEY, SHIPMENTTRANSFER.COMMENTMEMO "
            '    str2 = "FROM SHIPMENTTRANSFER "
            '    str3 = "WHERE(((SHIPMENTTRANSFER.TRANSFERKIND) = 1) And ((SHIPMENTTRANSFER.ACTIVESHIPMENT) = 'Y') And ((SHIPMENTTRANSFER.STUDYID) = " & wStudyID & ")) "
            '    str4 = "ORDER BY SHIPMENTTRANSFER.TRANSFERDATE, SHIPMENTTRANSFER.SHIPMENTTRANSFERKEY;"
            'Else
            '    str1 = "SELECT " & strSchema & ".SHIPMENTTRANSFER.SHIPMENTID, " & strSchema & ".SHIPMENTTRANSFER.STORAGETEMPERATURE, " & strSchema & ".SHIPMENTTRANSFER.SAMPLECONDITION, " & strSchema & ".SHIPMENTTRANSFER.TRANSFERDATE, " & strSchema & ".SHIPMENTTRANSFER.STUDYID, " & strSchema & ".SHIPMENTTRANSFER.SHIPMENTTRANSFERKEY, " & strSchema & ".SHIPMENTTRANSFER.COMMENTMEMO "
            '    str2 = "FROM " & strSchema & ".SHIPMENTTRANSFER "
            '    str3 = "WHERE(((" & strSchema & ".SHIPMENTTRANSFER.TRANSFERKIND) = 1) And ((" & strSchema & ".SHIPMENTTRANSFER.ACTIVESHIPMENT) = 'Y') And ((" & strSchema & ".SHIPMENTTRANSFER.STUDYID) = " & wStudyID & ")) "
            '    str4 = "ORDER BY " & strSchema & ".SHIPMENTTRANSFER.TRANSFERDATE, " & strSchema & ".SHIPMENTTRANSFER.SHIPMENTTRANSFERKEY;"
            'End If
            'strSQL = str1 & str2 & str3 & str4

            'Dim rsWShipment As New ADODB.Recordset

            'Try
            '    rsWShipment.CursorLocation = CursorLocationEnum.adUseClient
            '    rsWShipment.Open(strSQL, cn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
            '    rsWShipment.ActiveConnection = Nothing
            '    int1 = rsWShipment.RecordCount 'debug
            'Catch ex As Exception
            '    MsgBox("There was a problem opening rsWShipment." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            'End Try

            '''Console.WriteLine("rsWShipment: " & strSQL)

            'str1 = "Retrieving Watson Data...23d " & ctPB
            'str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            'frmH.lblProgress.Text = str1
            'frmH.pb1.Refresh()
            'frmH.lblProgress.Refresh()
            'System.Windows.Forms.Application.DoEvents()

            'Dim rsSamples As New ADODB.Recordset
            'If boolAccess Then
            '    str1 = "SELECT SHIPMENTSAMPLEID.* FROM SHIPMENTSAMPLEID "
            '    str2 = "WHERE(((SHIPMENTSAMPLEID.STUDYID) = " & wStudyID & "));"
            'Else
            '    str1 = "SELECT " & strSchema & ".SHIPMENTSAMPLEID.* FROM " & strSchema & ".SHIPMENTSAMPLEID "
            '    str2 = "WHERE(((" & strSchema & ".SHIPMENTSAMPLEID.STUDYID) = " & wStudyID & "));"
            'End If

            'strSQL = str1 & str2
            'If rsSamples.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    rsSamples.Close()
            'End If
            'Try
            '    rsSamples.CursorLocation = CursorLocationEnum.adUseClient
            '    rsSamples.Open(strSQL, cn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
            '    rsSamples.ActiveConnection = Nothing
            'Catch ex As Exception
            '    MsgBox("There was a problem opening rsSamples." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            'End Try

            'str1 = "Retrieving Watson Data...23e " & ctPB
            'str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            'frmH.lblProgress.Text = str1
            'frmH.pb1.Refresh()
            'frmH.lblProgress.Refresh()
            'System.Windows.Forms.Application.DoEvents()

            ''clear local tbl
            'tblSRecWatson.Clear()
            'Dim tbl As System.Data.DataTable
            'tbl = tblSRecWatson
            'int3 = 0 'for sample count sum
            'For Count1 = 0 To int1 - 1
            '    Dim drow1 As DataRow = tbl.NewRow
            '    drow1.BeginEdit()
            '    drow1.Item("Watson ID") = wStudyID
            '    drow1.Item("Date Received") = rsWShipment.Fields("TRANSFERDATE").Value
            '    drow1.Item("Storage Temperature") = rsWShipment.Fields("STORAGETEMPERATURE").Value
            '    drow1.Item("Sample Condition") = rsWShipment.Fields("SAMPLECONDITION").Value
            '    drow1.Item("STUDYID") = wStudyID
            '    drow1.Item("Comments") = rsWShipment.Fields("COMMENTMEMO").Value
            '    'find sample count
            '    var1 = rsWShipment.Fields("SHIPMENTTRANSFERKEY").Value
            '    str1 = "SHIPMENTTRANSFERKEY = " & var1
            '    rsSamples.Filter = ""
            '    rsSamples.Filter = str1
            '    int2 = rsSamples.RecordCount
            '    'sum Sample Count
            '    int3 = int3 + int2

            '    drow1.Item("Sample Count") = int2
            '    drow1.EndEdit()
            '    tbl.Rows.Add(drow1)

            '    rsWShipment.MoveNext()
            'Next


            ''total samplecount
            'Dim intSRC As Int32 = 0
            'Dim intSRCa As Int32 = 0
            'For Count1 = 0 To tblSampleReceipt.Rows.Count - 1
            '    intSRCa = NZ(tblSampleReceipt.Rows(Count1).Item("SAMPLECOUNT"), 0)
            '    intSRC = intSRC + intSRCa
            'Next

            'frmH.txtSRecTotalReportWatson.Text = intSRC
            'dv = tblSampleReceiptWatson.DefaultView
            'frmH.dgvSampleReceiptWatson.DataSource = dv
            'frmH.dgvSampleReceiptWatson.Refresh()
            'frmH.dgvSampleReceiptWatson.AutoResizeColumns()

            str1 = "Retrieving Watson Data...26 " & ctPB
            str1 = str1 & ChrW(10) & "...If the study is large, this step may take a few moments..."
            str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
            frmH.lblProgress.Text = str1

            Cursor.Current = Cursors.WaitCursor

            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            System.Windows.Forms.Application.DoEvents()

            'rsSamples.Close()
            'rsWShipment.Close()
            'rsSamples = Nothing
            'rsWShipment = Nothing

            'Hmm. The next command is filling from previous study
            Try
                Call FillAnalysisResultsTable(cn)
            Catch ex As Exception
                MsgBox("There was a problem executing FillAnalysisResultsTable." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            End Try



            Cursor.Current = Cursors.WaitCursor
            'add unbound columns to tblAssignedSamples
            Try
                Call AddCols_tblAss()
            Catch ex As Exception
                MsgBox("There was a problem executing AddCols_tblAss." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            End Try
            Cursor.Current = Cursors.WaitCursor

            'set dgvFC
            Try
                Call FillFCRW()
            Catch ex As Exception
                MsgBox("There was a problem executing FillFCRW." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            End Try
            Cursor.Current = Cursors.WaitCursor

            'pesky
            Try
                Call SampleReceiptChange()
            Catch ex As Exception
                MsgBox("There was a problem executing SampleReceiptChange." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            End Try
            Cursor.Current = Cursors.WaitCursor

        Catch ex As Exception

            str2 = "Hmmm. There was a problem retrieving information from the Watson database."
            str2 = str2 & ChrW(10) & ChrW(10) & "Please screen shot this entire window and provide to your StudyDoc Administrator."
            str2 = str2 & ChrW(10) & ChrW(10) & "Err: " & ex.Message & ChrW(10) & ChrW(10) & "str1 = " & str1
            str2 = str2 & ChrW(10) & ChrW(10) & frmH.lblProgress.Text


            MsgBox(str2, MsgBoxStyle.Information, "Problem...")

        End Try

        'hide some rows in dgvWatsonAnalRef
        Call HideWatsonRows()


        str1 = "Retrieving Watson Data...27 " & ctPB
        str1 = "Preparing study " & str_cbxStudy & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()

        Cursor.Current = Cursors.WaitCursor

end1:

        'Call MethValAutoCol()


        'frm.Visible = False

        rs = Nothing
        Try
            rsBCStds = Nothing
        Catch ex As Exception

        End Try

        rs1 = Nothing
        rs2 = Nothing
        rs3 = Nothing
        Try
            rs4 = Nothing
        Catch ex As Exception

        End Try

        Try
            rs20 = Nothing
        Catch ex As Exception

        End Try


        rsRS = Nothing
        cn = Nothing

        daDoPr.Dispose()

    End Function

End Module
