Option Compare Text

Module modHomeCode

    Function CheckStudyDocAnalytes1() As Boolean

        '20180321 LEE:
        'TRUE means need to re-establish
        'if CheckStudyDocAnalytes1 is true, then must follow up later in doprepare with CheckStudyDocAnalytes2

        CheckStudyDocAnalytes1 = False

        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim strM As String

        Dim strF As String
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim strS As String = "INTORDER ASC"
        Dim strCol As String

        Dim rowsSDA() As DataRow = TBLSTUDYDOCANALYTES.Select(strF, strS)

        Dim intRowsTAH As Short = tblAnalytesHome.Rows.Count
        Dim intRowsSDA As Short = rowsSDA.Length
        Dim boolBad As Boolean = False

        Dim var1, var2, var3, var4

        If intRowsSDA = 0 Then
            GoTo end1
        End If

        '20180612 LEE:
        'This logic is bad. If user re-orders, then re-opens study, matching logic is incorrect
        'this affects StudyDoc 3.1.18 and newer
        'instead filter for each ANALYTEDESCRIPTION
        Dim strAD As String
        Dim strFAD As String


        'check tblAnalytesHome with rowssda
        boolBad = False
        If intRowsTAH = intRowsSDA Then
            'check
            For Count1 = 1 To intRowsSDA

                '20180612 LEE:
                strAD = rowsSDA(Count1 - 1).Item("AnalyteDescription")
                strFAD = "AnalyteDescription = '" & CleanText(strAD) & "'"
                Dim rows1() As DataRow = tblAnalytesHome.Select(strFAD, "", DataViewRowState.CurrentRows)
                'now evaluate rows1
                If rows1.Length = 0 Then
                    boolBad = True
                    Exit For
                End If

                For Count2 = 1 To 5

                    Select Case Count2
                        Case 1
                            strCol = "AnalyteDescription"
                        Case 2
                            strCol = "AnalyteID"
                        Case 3
                            strCol = "ORIGINALANALYTEDESCRIPTION"
                        Case 4
                            strCol = "MATRIX"
                        Case 5
                            strCol = "CALIBRSET"
                           

                    End Select

                    var1 = NZ(rowsSDA(Count1 - 1).Item(strCol), 0)
                    var2 = NZ(rows1(0).Item(strCol), 0)

                    Select Case Count2
                        Case 5
                            '20181130 LEE:
                            'Instead, compare number of Calibr Levels
                            var3 = UBound(Split(var1.ToString, ",", -1, CompareMethod.Text)) + 1
                            var4 = UBound(Split(var2.ToString, ",", -1, CompareMethod.Text)) + 1

                            var1 = var3
                            var2 = var4
                        Case Else

                    End Select
                    

                    If var1 = var2 Then
                    Else
                        boolBad = True
                        Exit For
                    End If

                Next Count2

                If boolBad Then
                    Exit For
                End If

            Next Count1

            If boolBad Then
                For Count2 = 0 To rowsSDA.Length - 1
                    rowsSDA(Count2).Delete()
                Next
                CheckStudyDocAnalytes1 = True
            End If

        Else
            For Count2 = 0 To rowsSDA.Length - 1
                rowsSDA(Count2).Delete()
            Next
            CheckStudyDocAnalytes1 = True
        End If

        If CheckStudyDocAnalytes1 Then
            strM = "A change has been detected in Analyte assignment in StudyDoc (e.g. different Analyte/Matrix/CalibrCure combinations)."
            strM = strM & ChrW(10) & ChrW(10)
            strM = strM & "Please note that any Assigned Samples for this study will be lost."
            strM = strM & ChrW(10) & ChrW(10)
            strM = strM & "Please remember to re-Assign these samples."
            MsgBox(strM, vbInformation, "Analyte assignment change...")
        End If

end1:

    End Function

    Sub CheckStudyDocAnalytes2()

        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32

        Dim int1 As Int32
        Dim int2 As Int32

        Dim strF As String
        Dim strF1 As String
        Dim strS As String = "INTORDER ASC"
        Dim strCol As String

        Dim tblSDA As DataTable = TBLSTUDYDOCANALYTES
        Dim tblTAH As DataTable = tblAnalytesHome

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND IsIntStd = 'No'"
        Dim rowsSDA() As DataRow = tblSDA.Select(strF, strS)

        Dim intRowsTAH As Short = tblTAH.Rows.Count
        Dim intRowsSDA As Short = rowsSDA.Length
        Dim boolBad As Boolean = False

        Dim var1, var2, var3, var4, var5, var6, var7

        Dim boolHit As Boolean '20190219 LEE:

        Dim maxID As Int64

        Try

            'get unique intGroup from tblReportTableAnalytes
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            Dim dv1 As DataView = New DataView(tblReportTableAnalytes, strF, "ANALYTEID ASC, INTGROUP ASC", DataViewRowState.CurrentRows)
            Dim tblU As DataTable = dv1.ToTable("a", True, "INTGROUP", "ANALYTEID")
            Dim intU As Short = tblU.Rows.Count

            'compare tblu and tblsda
            Dim arrUGood(2, intRowsSDA)
            Dim intUGood As Short = 0
            Dim arrUBad(2, intRowsSDA)
            Dim intUBad As Short = 0

            Dim rows1() As DataRow

            For Count1 = 0 To intRowsSDA - 1
                var1 = rowsSDA(Count1).Item("ANALYTEID")
                var2 = NZ(rowsSDA(Count1).Item("INTGROUP"), -1)
                strF1 = "ANALYTEID = " & var1 & " AND INTGROUP = " & var2
                rows1 = tblTAH.Select(strF1, "", DataViewRowState.CurrentRows)
                If rows1.Length = 0 Then
                    intUBad = intUBad + 1
                    arrUBad(1, intUBad) = var1
                    arrUBad(2, intUBad) = var2
                Else
                    intUGood = intUGood + 1
                    arrUGood(1, intUGood) = var1
                    arrUGood(2, intUGood) = var2
                End If
            Next

            Dim boolGo As Boolean = False
            Dim intI As Short = 0
            Dim tbl2 As DataTable

            If intUGood = intU Then
            Else

                boolGo = False
                intI = 0
                Do Until boolGo
                    intI = intI + 1

                    If intI > intUGood Then
                        Exit Do
                    End If

                    If intUGood = 0 Then
                        'this should never happen
                    Else
                        var1 = arrUGood(1, intI) 'analyteid
                        var2 = arrUGood(2, intI) 'intgroup
                    End If


                    'strF1 = "ANALYTEID = " & var1 & " AND INTGROUP = " & var2 & " AND ID_TBLSTUDIES = " & id_tblStudies
                    'only need to look at group
                    strF1 = "INTGROUP = " & var2 & " AND ID_TBLSTUDIES = " & id_tblStudies
                    Dim rows2() As DataRow = tblReportTableAnalytes.Select(strF1, "ID_TBLREPORTTABLEANALYTES", DataViewRowState.CurrentRows)
                    int1 = rows2.Length
                    If int1 = 0 Then
                    Else
                        'save this info in a new table because data will get deleted next
                        'will use this info to create new records
                        tbl2 = rows2.CopyToDataTable
                        boolGo = True
                        Exit Do
                    End If

                Loop

                If boolGo Then
                Else
                    boolGo = False
                    intI = 0

                    Do Until boolGo
                        intI = intI + 1

                        If intI > intUBad Then
                            Exit Do
                        End If

                        If intUBad = 0 Then
                            'this should never happen
                        Else
                            var1 = arrUBad(1, intI) 'analyteid
                            var2 = arrUBad(2, intI) 'intgroup
                        End If


                        strF1 = "ANALYTEID = " & var1 & " AND INTGROUP = " & var2 & " AND ID_TBLSTUDIES = " & id_tblStudies
                        Dim rows2() As DataRow = tblReportTableAnalytes.Select(strF1, "ID_TBLREPORTTABLEANALYTES", DataViewRowState.CurrentRows)
                        int1 = rows2.Length
                        If int1 = 0 Then
                        Else
                            'save this info in a new table because data will get deleted next
                            'will use this info to create new records
                            tbl2 = rows2.CopyToDataTable
                            boolGo = True
                            Exit Do
                        End If

                    Loop

                End If


                'delete all data from tblReportTableAnalytes
                strF = "ID_TBLSTUDIES = " & id_tblStudies
                Dim rowsRTA() As DataRow = tblReportTableAnalytes.Select(strF, "", DataViewRowState.CurrentRows)
                Try
                    For Count2 = 0 To rowsRTA.Length - 1
                        rowsRTA(Count2).BeginEdit()
                        rowsRTA(Count2).Delete()
                        rowsRTA(Count2).EndEdit()
                    Next
                Catch ex As Exception
                    var1 = ex.Message
                End Try

                'now make new entries based on tbl2
                maxID = GetMaxID("tblReportTableAnalytes", 1, False)

                For Count3 = 0 To intRowsSDA - 1

                    var3 = rowsSDA(Count3).Item("ANALYTEID")
                    var4 = rowsSDA(Count3).Item("ANALYTEINDEX")
                    var5 = rowsSDA(Count3).Item("INTGROUP")
                    var6 = rowsSDA(Count3).Item("MASTERASSAYID")

                    For Count1 = 0 To tbl2.Rows.Count - 1

                        maxID = maxID + 1
                        boolhit = True

                        Dim nr As DataRow = tblReportTableAnalytes.NewRow
                        nr.BeginEdit()

                        For Count2 = 0 To tbl2.Columns.Count - 1

                            var1 = tbl2.Columns(Count2).ColumnName
                            Select Case var1
                                Case "ID_TBLREPORTTABLEANALYTES"
                                    var2 = maxID
                                Case "ANALYTEID"
                                    var2 = var3
                                Case "ANALYTEINDEX"
                                    var2 = var4
                                Case "INTGROUP"
                                    var2 = var5
                                Case "MASTERASSAYID"
                                    var2 = var6
                                Case Else
                                    var2 = tbl2.Rows(Count1).Item(Count2)
                            End Select

                            nr.Item(Count2) = var2

                        Next Count2

                        nr.EndEdit()

                        tblReportTableAnalytes.Rows.Add(nr)

                    Next Count1

                Next

                'save results
                If boolGuWuOracle Then
                    Try
                        ta_tblReportTableAnalytes.Update(tblReportTableAnalytes)
                    Catch ex As DBConcurrencyException
                        ds2005.TBLREPORTTABLEANALYTES.Merge(ds2005.TBLREPORTTABLEANALYTES, True)
                    End Try
                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblReportTableAnalytesAcc.Update(tblReportTableAnalytes)
                    Catch ex As DBConcurrencyException
                        ds2005Acc.TBLREPORTTABLEANALYTES.Merge(ds2005Acc.TBLREPORTTABLEANALYTES, True)
                    End Try
                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblReportTableAnalytesSQLServer.Update(tblReportTableAnalytes)
                    Catch ex As DBConcurrencyException
                        ds2005Acc.TBLREPORTTABLEANALYTES.Merge(ds2005Acc.TBLREPORTTABLEANALYTES, True)
                    End Try
                End If

            End If

            If boolhit Then
                Call PutMaxID("tblReportTableAnalytes", maxID)
            End If


            'delete entries from tblAssignedSamples
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            Dim rowsAS() As DataRow = tblAssignedSamples.Select(strF, "", DataViewRowState.CurrentRows)
            int1 = rowsAS.Length
            int2 = 0
            For Count2 = 0 To rowsAS.Length - 1
                int2 = int2 + 1
                rowsAS(Count2).BeginEdit()
                rowsAS(Count2).Delete()
                rowsAS(Count2).EndEdit()
            Next
            'now must save tblassignedsamples


            If rowsAS.Length = 0 Then
            Else
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
            End If

        Catch ex As Exception
            var1 = ex.Message
        End Try

        GoTo end1

end1:

    End Sub

    Sub HideWatsonRows()

        Exit Sub

        Dim dgv As DataGridView = frmH.dgvWatsonAnalRef

        Dim intRows As Short = dgv.Rows.Count
        Dim Count1 As Short

        For Count1 = 3 To intRows - 1

            dgv.Rows(Count1).Visible = False

        Next

    End Sub

    Function RecordAuditTrail(ByVal boolUseDt As Boolean, ByVal dtR As DateTime) As DateTime

        Dim dtAT As DateTime

        If boolUseDt Then
            dtAT = dtR
        Else
            dtAT = Now
        End If

        RecordAuditTrail = dtAT

        'audit trail stuff
        If gboolAuditTrail Then
        Else
            Exit Function
        End If

        If idSE = 0 Then 'nothing to save
        Else

            Dim id1 As Double
            Dim id2 As Double
            Dim id3 As Double
            Dim idM As Double
            Dim strCol As String
            Dim Count1 As Long
            Dim Count2 As Long
            Dim colAT As New System.Data.DataColumn

            '****

            Const dataFmt As String = "{0,-30}{1}"
            'Const timeFmt As String = "{0,-30}{1:yyyy-MM-dd HH:mm}"
            Const timeFmt As String = "{0,-30}{1:yyyy-MMM-dd HH:mm}"

            ''''''''console.writeline("This example of selected " & _
            '    "TimeZone class elements generates the following " & _
            '    vbCrLf & "output, which varies depending on the " & _
            '    "time zone in which it is run." & vbCrLf)

            ' Get the local time zone and the current local time and year.
            Dim localZone As TimeZone = TimeZone.CurrentTimeZone
            Dim currentDate As DateTime = dtAT 'DateTime.Now
            Dim currentYear As Integer = currentDate.Year

            'record tblsaveevent
            'this records gSID
            Call RecordSaveEvent(currentDate)

            Dim strTimeZoneName As String
            Dim strDaylightName As String
            Dim boolDST As Boolean = False
            Dim strCUT As String
            Dim strOffset As String

            ' Display the names for standard time and daylight saving 
            ' time for the local time zone.
            ''''''''console.writeline(dataFmt, "Standard time name:", localZone.StandardName)
            strTimeZoneName = NZ(localZone.StandardName, "NA")
            ''''''''console.writeline(dataFmt, "Daylight saving time name:", localZone.DaylightName)
            strDaylightName = NZ(localZone.DaylightName, "NA")

            ' Display the current date and time and show if they occur 
            ' in daylight saving time.
            ''''''''console.writeline(vbCrLf & timeFmt, "Current date and time:", currentDate)
            ''''''''console.writeline(dataFmt, "Daylight saving time?", localZone.IsDaylightSavingTime(currentDate))
            boolDST = localZone.IsDaylightSavingTime(currentDate)
            ' Get the current Coordinated Universal Time (UTC) and UTC 
            ' offset.
            Dim currentUTC As DateTime = localZone.ToUniversalTime(currentDate)
            Dim currentOffset As TimeSpan = localZone.GetUtcOffset(currentDate)

            strCUT = Format(currentUTC, "MMM dd, yyyy HH:mm:ss tt")
            strOffset = currentOffset.ToString

            ''''''''console.writeline(timeFmt, "Coordinated Universal Time:", currentUTC)
            ''''''''console.writeline(dataFmt, "UTC offset:", currentOffset)

            '' Get the DaylightTime object for the current year.
            'Dim daylight As DaylightTime = _
            '    localZone.GetDaylightChanges(currentYear)

            '' Display the daylight saving time range for the current year.
            ''''''''console.writeline(vbCrLf & _
            '    "Daylight saving time for year {0}:", currentYear)
            ''''''''console.writeline("{0:yyyy-MM-dd HH:mm} to " & _
            '    "{1:yyyy-MM-dd HH:mm}, delta: {2}", _
            '    daylight.Start, daylight.End, daylight.Delta)

            '***


            id1 = GetMaxID("TBLAUDITTRAIL", 1, True)
            idM = id1
            id2 = id1 + idSE
            id3 = id2 + 1
            Call PutMaxID("TBLAUDITTRAIL", CLng(id3))
            For Count1 = 0 To tblAuditTrailTemp.Rows.Count - 1

                Dim nr As DataRow = tblAuditTrail.NewRow
                nr.BeginEdit()

                idM = idM + 1

                nr("ID_TBLAUDITTRAIL") = idM

                For Each colAT In tblAuditTrailTemp.Columns
                    strCol = colAT.ColumnName
                    If StrComp(strCol, "ID_TBLAUDITTRAIL", CompareMethod.Text) = 0 Then 'ignore
                    Else
                        nr(strCol) = tblAuditTrailTemp.Rows(Count1).Item(strCol)
                    End If
                Next

                nr("ID_TBLSAVEEVENT") = gSID
                nr("DTSAVEDATE") = currentDate ' dtAT
                nr("CHARSTANDARDTIMEZONE") = strTimeZoneName
                nr("CHARDAYLIGHTSAVINGZONE") = strDaylightName
                nr("CHARDAYLIGHTSAVINGTIME") = boolDST.ToString
                nr("CHARCOORUNIVTIME") = strCUT
                nr("DTCOORUNIVTIME") = currentUTC
                nr("CHARUTCOFFSET") = strOffset
                nr("CHARTBLREASONFORCHANGE") = strRFC
                nr("CHARTBLCHARMEANINGOFSIG") = strMOS

                If gboolLDAP Then
                    nr("CHARUSERID") = gUserID & " (Logged in from Network User ID " & gNetAcct & ")"
                Else
                    nr("CHARUSERID") = gUserID
                End If

                nr.EndEdit()

                tblAuditTrail.Rows.Add(nr)

            Next

            gSID = 0

            Call SaveAuditTrail()


        End If

    End Function


    Sub RecordSaveEvent(ByVal dt1 As DateTime)

        Dim maxID As Int64
        maxID = GetMaxID("TBLSAVEEVENT", 1, True) 'this already incremented

        gSID = maxID

        Dim nr1 As DataRow = tblSaveEvent.NewRow
        nr1.BeginEdit()
        nr1.Item("ID_TBLSAVEEVENT") = maxID
        nr1.Item("DTSAVEDATE") = dt1
        nr1.Item("CHARUSERNAME") = gUserName
        nr1.Item("CHARUSERID") = gUserID
        nr1.Item("ID_TBLREASONFORCHANGE") = 0 'DON'T NEED THIS
        nr1.Item("ID_TBLCHARMEANINGOFSIG") = 0 'DON'T NEED THIS

        nr1.Item("CHARREASONFORCHANGE") = strRFC
        nr1.Item("CHARMEANINGOFSIG") = strMOS


        'legend
        'strAuditType = "Report Writer Study"
        'strAuditType = "StudyDoc Administration"
        'strAuditType = "Report Writer Administration"
        'strAuditType = "Study Design"
        'strAuditType = "Study Design Administration"

        If InStr(1, strAuditType, "Administration", CompareMethod.Text) > 0 Then
            nr1("ID_TBLSTUDIES") = -1
        Else
            nr1("ID_TBLSTUDIES") = id_tblStudies
        End If

        If InStr(1, strAuditType, "Report Writer Study", CompareMethod.Text) > 0 Then
            nr1.Item("BOOLREPORTWRITER") = -1
        Else
            nr1.Item("BOOLREPORTWRITER") = 0
        End If

        If InStr(1, strAuditType, "StudyDoc Admin", CompareMethod.Text) > 0 Then
            nr1.Item("BOOLGUWUADMIN") = -1
        Else
            nr1.Item("BOOLGUWUADMIN") = 0
        End If

        If InStr(1, strAuditType, "Report Writer Admin", CompareMethod.Text) > 0 Then
            nr1.Item("BOOLREPORTWRITERADMIN") = -1
        Else
            nr1.Item("BOOLREPORTWRITERADMIN") = 0
        End If

        If InStr(1, strAuditType, "Study Design Study", CompareMethod.Text) > 0 Then
            nr1.Item("BOOLSTUDYDESIGN") = -1
        Else
            nr1.Item("BOOLSTUDYDESIGN") = 0
        End If

        If InStr(1, strAuditType, "Study Design Admin", CompareMethod.Text) > 0 Then
            nr1.Item("BOOLSTUDYDESIGNADMIN") = -1
        Else
            nr1.Item("BOOLSTUDYDESIGNADMIN") = 0
        End If

        nr1.Item("INTNUMADDS") = gATAdds
        nr1.Item("INTNUMDELETES") = gATDeletes
        nr1.Item("INTNUMMODIFIES") = gATMods

        nr1.EndEdit()

        tblSaveEvent.Rows.Add(nr1)

        Call PutMaxID("TBLSAVEEVENT", maxID)

        If boolGuWuOracle Then
            Try
                ta_tblSaveEvent.Update(tblSaveEvent)
            Catch ex As DBConcurrencyException
                'ds2005.tblSaveEvent.Merge('ds2005.tblSaveEvent, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblSaveEventAcc.Update(tblSaveEvent)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLSAVEEVENT.Merge('ds2005Acc.TBLSAVEEVENT, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblSaveEventSQLServer.Update(tblSaveEvent)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLSAVEEVENT.Merge('ds2005Acc.TBLSAVEEVENT, True)
            End Try
        End If

        gATAdds = 0
        gATDeletes = 0
        gATMods = 0

    End Sub


    Sub SaveAuditTrail()

        Dim var1

        Try
            If boolGuWuOracle Then
                Try
                    ta_tblAuditTrail.Update(tblAuditTrail)
                Catch ex As DBConcurrencyException
                    'ds2005.tblAuditTrail.Merge('ds2005.tblAuditTrail, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblAuditTrailAcc.Update(tblAuditTrail)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLAUDITTRAIL.Merge('ds2005Acc.TBLAUDITTRAIL, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblAuditTrailSQLServer.Update(tblAuditTrail)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLAUDITTRAIL.Merge('ds2005Acc.TBLAUDITTRAIL, True)
                End Try
            End If
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try



    End Sub


    Sub CheckWatsonRecords()

        Dim str1 As String = CheckWatsonRecordsString()

        frmH.lblWarning.Text = str1

        If InStr(1, str1, "Warning", CompareMethod.Text) > 0 Then
            frmH.TimerWarning.Enabled = True
            frmH.lblWarning.BackColor = System.Drawing.Color.FromArgb(255, 240, 240)
            frmH.lblWarning.ForeColor = Color.Red
        Else
            frmH.TimerWarning.Enabled = False
            frmH.lblWarning.BackColor = Color.White
            frmH.lblWarning.ForeColor = System.Drawing.Color.FromArgb(49, 112, 193)
            frmH.panDot.Visible = False
        End If

    End Sub

    Function ReturnWatsonCheckColumn() As String

        ReturnWatsonCheckColumn = ""

        Dim Count1 As Short

        '*****
        Dim var1, var2
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String

        Dim rows() As DataRow
        Dim rowsFR() As DataRow
        Dim dtR As Date
        Dim dtFR As Date
        Dim boolFR As Boolean = False
        Dim boolUseRH As Boolean = False 'use tblReportHistory
        Dim boolUseFR As Boolean = False 'use tblFinalReport and gboolER
        Dim boolNothing As Boolean = False
        Dim strCol As String

        If gboolER Then
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARREPORTTYPE LIKE '*FINAL*'"
            strS = "UPSIZE_TS DESC"
            rows = tblFinalReport.Select(strF, strS)
            If rows.Length = 0 Then
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARREPORTGENERATEDSTATUS LIKE '*FINAL*'"
                strS = "DTREPORTGENERATED DESC"
                rows = tblReportHistory.Select(strF, strS)
                If rows.Length = 0 Then
                    boolNothing = True
                Else
                    boolUseRH = True
                End If
            Else
                boolUseFR = True
            End If

        Else

            boolUseFR = False
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARREPORTGENERATEDSTATUS LIKE '*FINAL*'"
            strS = "DTREPORTGENERATED DESC"
            rows = tblReportHistory.Select(strF, strS)
            If rows.Length = 0 Then
                boolNothing = True
            Else
                boolUseRH = True
            End If

        End If

        If boolNothing Then
            GoTo end1
        End If

        If boolUseFR Then 'this is gboolER
            strCol = "UPSIZE_TS"
        Else
            strCol = "DTREPORTGENERATED"
        End If

        var1 = rows(0).Item(strCol)
        If IsDBNull(var1) Or IsDate(var1) = False Then
            GoTo end1
        End If

        ReturnWatsonCheckColumn = strCol
end1:

    End Function

    Function GetWatsonWarningRecordsNumber(dtR As Date) As Int32

        GetWatsonWarningRecordsNumber = 0

        Dim Count1 As Short

        '*****
        Dim var1, var2
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String

        Dim rows() As DataRow
        Dim rowsFR() As DataRow
        Dim dtFR As Date
        Dim boolFR As Boolean = False
        Dim boolUseRH As Boolean = False 'use tblReportHistory
        Dim boolUseFR As Boolean = False 'use tblFinalReport and gboolER
        Dim boolNothing As Boolean = False


        'get analytes
        Dim intAH As Int16 = tblAnalytesHome.Rows.Count
        Dim intAH1 As Int16 = 0
        For Count1 = 0 To intAH - 1
            var1 = NZ(tblAnalytesHome.Rows(Count1).Item("ANALYTEDESCRIPTION"), "")
            var2 = NZ(tblAnalytesHome.Rows(Count1).Item("ISINTSTD"), "Yes")
            If Len(var1) = 0 Then
            Else
                If StrComp(var2, "No", CompareMethod.Text) = 0 Then
                    intAH1 = intAH1 + 1
                    If intAH1 = 1 Then
                        strF1 = "(CHARANALYTE = '" & CleanText(CStr(var1)) & "')"
                    Else
                        strF1 = strF1 & " OR (CHARANALYTE = '" & CleanText(CStr(var1)) & "')"
                    End If
                End If
            End If
        Next

        'recordtimestamp comes from ANARUNRAWANALYTEPEAK
        strF = "RECORDTIMESTAMP > '" & Format(dtR, "dd-MMM-yyyy hh:mm tt") & "'"
        If intAH1 = 0 Then
            strF2 = strF
        Else
            strF2 = strF & " AND (" & strF1 & ")"
        End If

        'str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"
        strS = "CHARANALYTE ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

        Dim dv1 As DataView
        Try
            dv1 = New DataView(tblAnalysisResultsHome, strF2, strS, DataViewRowState.CurrentRows)
            Dim intRR As Int64
            intRR = dv1.Count
            intRR = intRR

            GetWatsonWarningRecordsNumber = intRR
        Catch ex As Exception
            var1 = ex.Message
            GetWatsonWarningRecordsNumber = 0
        End Try
      

    End Function

    Function GetWatsonErrorStatement(intRR As Int32, boolUseFR As Boolean, dtFR As Date) As String

        GetWatsonErrorStatement = ""

        'Legend
        'Dim boolUseRH As Boolean = False 'use tblReportHistory
        'Dim boolUseFR As Boolean = False 'use tblFinalReport and gboolER

        Dim str1 As String
        Dim str2 As String

        If intRR = 0 Then
            str1 = "The last Final Report saved is current." & ChrW(10) & ChrW(10)
            If boolUseFR Then
                str1 = str1 & "No Watson samples have been modified since the last Final Report was saved:" & ChrW(10) & Format(dtFR, "dd-MMM-yyyy hh:mm tt")
            Else
                str1 = str1 & "No Watson samples have been modified since the last Final Report was generated:" & ChrW(10) & Format(dtFR, "dd-MMM-yyyy hh:mm tt")
            End If

            boolWatsonWarning = False

        Else
            If boolUseFR Then
                str1 = "WARNING!" & ChrW(10) & ChrW(10) & intRR & " Watson samples have been modified since the last Final Report was saved" & ChrW(10) & "(" & Format(dtFR, "dd-MMM-yyyy hh:mm tt") & ")."
            Else
                str1 = "WARNING!" & ChrW(10) & ChrW(10) & intRR & " Watson samples have been modified since the last Final Report was generated" & ChrW(10) & "(" & Format(dtFR, "dd-MMM-yyyy hh:mm tt") & ")."
            End If
            If intRR = 1 Then
                str2 = Replace(str1, "samples have been", "sample has been", 1, -1, CompareMethod.Text)
                str1 = str2
            End If
            str1 = str1 & ChrW(10) & ChrW(10) & "Click 'View Report History' button for further details."

            boolWatsonWarning = True

        End If

        GetWatsonErrorStatement = str1

    End Function

    Function CheckWatsonRecordsString() As String

        Dim Count1 As Short

        '*****
        Dim var1, var2
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String

        Dim rows() As DataRow
        Dim rowsFR() As DataRow
        Dim dtR As Date
        Dim dtFR As Date
        Dim boolFR As Boolean = False
        Dim boolUseRH As Boolean = False 'use tblReportHistory
        Dim boolUseFR As Boolean = False 'use tblFinalReport and gboolER
        Dim boolNothing As Boolean = False
        Dim strCol As String
        Dim strCol1 As String
        Dim tbl1 As DataTable

        strCol = ReturnWatsonCheckColumn()

        If Len(strCol) = 0 Then
            boolWatsonWarning = False
            str1 = "A Final Report has not been generated for this study."
            CheckWatsonRecordsString = str1
            frmH.lblWarning.Text = str1
            frmH.TimerWarning.Enabled = False
            frmH.panDot.Visible = False
            GoTo end1
        End If

        'set boolusefr
        Select Case strCol
            Case "UPSIZE_TS"
                boolUseFR = True
                strCol1 = "CHARREPORTTYPE"
                tbl1 = tblFinalReport
            Case "DTREPORTGENERATED"
                boolUseRH = True
                strCol1 = "CHARREPORTGENERATEDSTATUS"
                tbl1 = tblReportHistory
            Case Else
                boolUseRH = True
        End Select

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND " & strCol1 & " LIKE '*FINAL*'"
        strS = strCol & " DESC"
        rows = tbl1.Select(strF, strS)

        'Note that ReturnWatsonCheckColumn already checks to see if rows(0) contents is OK 
        var1 = rows(0).Item(strCol)
        dtR = CDate(var1)
        dtFR = dtR
        gWatsonCutOffDt = dtR

        'now look for dates in Watson data
        Dim intRR As Int64
        intRR = GetWatsonWarningRecordsNumber(dtR)


        'Note that GetWatsonErrorStatement sets the global variable boolWatsonWarning
        CheckWatsonRecordsString = GetWatsonErrorStatement(intRR, boolUseFR, dtR)

end1:


    End Function

    Sub SetFormPos(frm As Form)

        frm.WindowState = FormWindowState.Maximized

        Exit Sub

        Dim w, h
        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        'Me.Top = 0
        'Me.Left = 0
        'Me.Width = w
        'Me.Height = h

        frm.Top = h * 0.05
        frm.Left = w * 0.05
        frm.Width = w * 0.9
        frm.Height = h * 0.9

    End Sub

    Sub PositionProgress()

        Dim int1 As Short

        int1 = frmH.lbxTab1.SelectedIndex

        Dim x1, x2, x3, b1, b2, tp1, tp2, tp3

        'Select Case int1
        '    Case 5 'Report tables
        '        frmH.lblProgress.Size = frmH.dgvReportTableConfiguration.Size
        '    Case Else
        '        frmH.lblProgress.Size = frmH.dgvwStudy.Size
        'End Select

        'frmH.lblProgress.Size = frmH.tp1.Size ' frmH.dgvwStudy.Size
        'frmH.panProgress.Size = frmH.tab1.Size ' frmH.tp1.Size

        'lblProgress.Size = dgvwStudy.Size
        x1 = frmH.tab1.Left
        Select Case int1
            Case 5 'Tables
                x2 = frmH.dgvReportTableConfiguration.Left
            Case Else
                x2 = frmH.dgvwStudy.Left
        End Select
        x3 = frmH.tp1.Left
        b1 = 4
        'b2 = frmH.pb1.Height * 5 / 100 ' 20

        'new
        'frmH.lblProgress.Left = x2 + b1
        'tp1 = frmH.tab1.Top
        'tp2 = frmH.dgvwStudy.Top
        'frmH.lblProgress.Top = tp2 + b1

        'old
        'frmH.lblProgress.Left = x1 + x3 ' x1 + x2 + x3
        tp1 = frmH.tab1.Top
        tp2 = frmH.tp1.Top
        'Select Case int1
        '    Case 5 'Tables
        '        tp3 = frmH.dgvReportTableConfiguration.Top
        '    Case Else
        '        tp3 = frmH.dgvwStudy.Top
        'End Select

        'frmH.lblProgress.Top = tp1 + tp2 ' + tp3
        'frmH.panProgress.Left = frmH.tab1.Left + frmH.tp1.Left
        frmH.panProgress.Left = frmH.lbxTab1.Left
        frmH.panProgress.Top = tp1 ' + tp2
        frmH.panProgress.Width = frmH.tab1.Left + frmH.tab1.Width - frmH.lbxTab1.Left
        frmH.panProgress.Height = frmH.tab1.Height

        Exit Sub

        b2 = 1
        'frmH.pb1.Top = frmH.lblProgress.Top + frmH.lblProgress.Height + b2
        frmH.pb1.Top = frmH.lblProgress.Top + frmH.lblProgress.Height - frmH.pb1.Height - frmH.pb2.Height - (b2 * 2)
        frmH.pb1.Left = frmH.lblProgress.Left
        frmH.pb1.Width = frmH.lblProgress.Width

        frmH.pb2.Top = frmH.pb1.Top + frmH.pb1.Height + b2
        frmH.pb2.Left = frmH.pb1.Left
        frmH.pb2.Width = frmH.pb1.Width

        frmH.lblProgress.BringToFront()
        frmH.pb1.BringToFront()
        frmH.pb2.BringToFront()


        'frmH.lblProgress.Refresh()
        'frmH.Refresh()

    End Sub

    Sub LockSections(wd As Microsoft.Office.Interop.Word.Application)

        Dim Count1 As Integer
        Dim Count2 As Integer
        Dim doc As Microsoft.Office.Interop.Word.Document
        Dim posL As Int64
        Dim posUL As Int64
        Dim posF As Int64
        Dim str1 As String
        Dim strM As String
        Dim boolLockStart As Boolean
        Dim boolLockEnd As Boolean
        Dim var1, var2

        boolLockStart = False
        boolLockEnd = False

        strM = "No Lock/Unlock pairs"

        Dim int1 As Int64
        Dim int2 As Int64
        Dim intLL As Integer
        Dim strPW As String

        strPW = "2@StudyDoc"
        strPW = ""

        'Int ((6 - 1 + 1) * Rnd + 1) would return a random number between 1 and 6
        'Int ((200 - 150 + 1) * Rnd + 150) would return a random number between 150 and 200
        'Int ((999 - 100 + 1) * Rnd + 100) would return a random number between 100 and 999
        'Int ((122 - 48 + 1) * Rnd + 48) would return a random number between 48 and 122
        'For Count1 = 1 To 16
        '    var1 = Int((122 - 48 + 1) * Rnd() + 48)
        '    var2 = ChrW(var1)
        '    strPW = strPW & var2
        'Next

        strPW = RandomPswd()
        'encrypt this passward
        tPswd = PasswordEncrypt(strPW) ' Decode(Coding(strPW, True), False) 'to be used when saving document

        doc = wd.ActiveDocument

        Dim arrL(2, 100)
        '1=UNLOCK Pos, 2=LOCK Pos
        Dim intL As Integer
        Dim intUL As Integer

        Dim strL As String
        Dim strUL As String

        strL = "[LOCKSECTION]"
        strUL = "[UNLOCKSECTION]"

        Dim rng As Microsoft.Office.Interop.Word.Range

        With wd


            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            rng = doc.Content

            Dim posO As Int64

            posO = .Selection.Start
            posF = posO + 1

            'first do some checking
            'find out how many LOCKS and UNLOCKS there are
            intL = 0
            Do Until posF = posO
                rng.Find.Execute(FindText:=strL, Forward:=True)
                If rng.Find.Found = True Then
                    intL = intL + 1
                    If intL = 1 Then
                        posL = rng.Start
                    End If
                Else
                    posF = posO
                End If
            Loop
            If intL = 0 Then
                GoTo end1
            End If

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng = doc.Content
            posO = .Selection.Start
            posF = posO + 1
            intLL = 0
            Do Until posF = posO
                rng.Find.Execute(FindText:=strUL, Forward:=True)
                If rng.Find.Found = True Then
                    intUL = intUL + 1
                    If intUL = 1 Then
                        posUL = rng.Start
                    End If
                Else
                    posF = posO
                End If
            Loop

            'evaluate
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng = doc.Content
            'If intL = 1 And intUL = 0 Then
            If intUL = 0 Then
                If posL = 0 Then
                    'protect entire document
                Else
                    int1 = 0
                    int2 = posL
                    rng = doc.Range(int1, int2)
                    rng.Select()

                    .Selection.Editors.Add(Microsoft.Office.Interop.Word.WdEditorType.wdEditorEveryone)
                    doc.Windows(1).View.ShadeEditableRanges = False
                End If
                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                'doc.Protect Password:=strPW, NoReset:=False, Type:=wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False
                GoTo end2
            End If


            'Continue

            'Home is first unlocked section

            'get locked
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng = doc.Content
            posO = .Selection.Start
            posF = posO + 1


            intL = 0
            Do Until posF = posO

                rng.Find.Execute(FindText:=strL, Forward:=True)
                If rng.Find.Found = True Then
                    posL = rng.Start
                    intL = intL + 1
                    If intL > UBound(arrL, 2) Then
                        ReDim Preserve arrL(2, UBound(arrL, 2) + 100)
                    End If
                    If intL = 1 Then
                        arrL(1, intL) = 0
                    End If
                    arrL(2, intL) = posL
                Else
                    posF = posO
                End If

            Loop
            If intL = 0 Then
                GoTo end1
            End If

            'get unlocked
            intUL = 1
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng = doc.Content
            posF = posO + 1

            posL = .Selection.Start
            Do Until posF = posO

                rng.Find.Execute(FindText:=strUL, Forward:=True)
                If rng.Find.Found = True Then
                    posL = rng.Start
                    intUL = intUL + 1
                    If intL > UBound(arrL, 2) Then
                        ReDim Preserve arrL(2, UBound(arrL, 2) + 100)
                    End If
                    arrL(1, intUL) = posL
                Else
                    If intUL = 0 Then
                    Else
                        posF = posO
                    End If
                End If

            Loop

            'add a final locked
            If intUL > intL Then 'make last part of document
                posL = doc.Range.End ' - 1 'end is paragraph return, which cannot be chosen
                arrL(2, intUL) = posL
                intL = intUL
            End If
            'arrL(2, intUL) = arrL(1, intUL)
            'intL = intL + 1

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            'use the lowest setting
            If intL = intUL Then
                intLL = intL
            ElseIf intL < intUL Then
                intLL = intL
            ElseIf intUL < intL Then
                intLL = intUL
            End If

            ReDim Preserve arrL(2, intLL)

            'now lock sections
            For Count1 = 1 To intLL

                int1 = arrL(1, Count1)
                int2 = arrL(2, Count1)
                rng = doc.Range(int1, int2)
                rng.Select()

                .Selection.Editors.Add(Microsoft.Office.Interop.Word.WdEditorType.wdEditorEveryone)

            Next

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

end2:
            'now erase all LOCKS and UNLOCKS
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng = doc.Content
            rng.Find.Execute(FindText:=strL, ReplaceWith:="", Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            rng = doc.Content
            rng.Find.Execute(FindText:=strUL, ReplaceWith:="", Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

            doc.Protect(Password:=strPW, NoReset:=False, Type:=Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False)
            doc.Windows(1).View.ShadeEditableRanges = True


            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        End With


        'debugging
        'For Count1 = 1 To intLL 'debugging
        '    str1 = "Lock: " & arrL(1, Count1) & " - Unlock: " & arrL(2, Count1)
        '    strM = strM & ChrW(10) & str1
        'Next

end1:

        'MsgBox strM

    End Sub

    Sub CheckforDoGuWu(wd As Microsoft.Office.Interop.Word.Application, strP As String)

        'find Name
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim intHit As Short
        Dim boolHit As Boolean
        Dim strName As String
        Dim strFind As String
        Dim boolFound As Boolean
        Dim var1, var2, var3
        Dim dtbl As System.Data.DataTable
        Dim var8
        Dim varReplace
        Dim tblNick As System.Data.DataTable
        Dim str2 As String
        Dim rowsNick() As DataRow
        Dim drows() As DataRow
        Dim dv As System.Data.DataView
        Dim strR As String = ""

        'find DoGuWu

        Dim mySel As Microsoft.Office.Interop.Word.Selection
        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
        wd.Selection.Find.ClearFormatting()

        mySel = wd.Selection
        strFind = "DoGuWu"
        boolFound = False
        With mySel.Find
            .ClearFormatting()
            '.Text = strFind
            .Forward = True
            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
            .Execute(FindText:=strFind)

            If .Found Then
                boolFound = True
                strR = strFind
            Else
                boolFound = False
            End If

        End With

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        If boolFound Then
        Else
            'look for DoStudyDoc
            strFind = "DoStudyDoc"
            boolFound = False
            With mySel.Find
                .ClearFormatting()
                '.Text = strFind
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                .Execute(FindText:=strFind)

                If .Found Then
                    boolFound = True
                    strR = strFind
                Else
                    boolFound = False
                End If

            End With

        End If

        If boolFound Then
        Else
            GoTo end2
        End If

        'determine if info tables are present
        Dim boolI As Boolean = False
        strFind = "DoGuWu Info"
        With mySel.Find
            .ClearFormatting()
            '.Text = strFind
            .Forward = True
            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
            .Execute(FindText:=strFind)

            If .Found Then
                boolI = True
            Else
                boolI = False
            End If

        End With

        If boolI = False Then 'keep looking
            strFind = "DoGuWuFast Info"
            With mySel.Find
                .ClearFormatting()
                '.Text = strFind
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                .Execute(FindText:=strFind)

                If .Found Then
                    boolI = True
                Else
                    boolI = False
                End If

            End With
        End If

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        If boolI Then 'complete rest of stuff
        Else

            'determine if info tables are present
            strFind = "DoStudyDoc Info"
            With mySel.Find
                .ClearFormatting()
                '.Text = strFind
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                .Execute(FindText:=strFind)

                If .Found Then
                    boolI = True
                Else
                    boolI = False
                End If

            End With

            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)


            If boolI = False Then 'keep looking
                strFind = "DoStudyDocFast Info"
                With mySel.Find
                    .ClearFormatting()
                    '.Text = strFind
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Execute(FindText:=strFind)

                    If .Found Then
                        boolI = True
                    Else
                        boolI = False
                    End If

                End With
            End If

            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        End If

        If boolI Then
        Else
            GoTo end2
        End If

        'start entering information

        Dim tbl As Microsoft.Office.Interop.Word.Table


        If wd.ActiveDocument.Tables.Count = 0 Then
            GoTo end2
        End If

        str1 = "Executing custom report..."
        str1 = str1 & ChrW(10) & ChrW(10) & "(A Word icon in the taskbar may flash a few times)"
        frmH.lblProgress.Text = str1
        frmH.lblProgress.Refresh()

        tbl = wd.ActiveDocument.Tables(1)

        'debug
        'Dim dgvW As DataGridView
        'dgvW = frmH.dgvwStudy
        'Dim intEE As Short
        'intEE = dgvW.ColumnCount
        'For Count1 = 0 To intEE - 1
        '    str1 = dgvW.Columns(Count1).Name
        '    str1 = str1

        'Next

        Dim tblRows As Short
        Dim intRows As Short
        Dim intCols As Short
        Dim tblCols As Short

        tblRows = tbl.Rows.Count
        intRows = 12 ' 10

        'wd.Visible = True

        'this doesn't work
        'If intRows > tblRows Then 'add rows
        '    tbl.Cell(tblRows, 1).Select()
        '    wd.Selection.InsertRowsBelow(tblRows - intRows)
        'End If

        With wd

            Dim boolDo As Boolean
            For Count1 = 1 To tblRows ' intRows
                boolDo = True
                Select Case Count1
                    Case 1
                        str1 = "Watson Study ID:"
                        var1 = wStudyID ' ""
                    Case 2
                        str1 = "GuWu Connection String:"
                        var1 = constrIni
                    Case 3
                        str1 = "Report Type"
                        var1 = ""
                    Case 4
                        str1 = "Table Field Code"
                        var1 = ""
                    Case 5
                        str1 = "Watson Connection String:"
                        var1 = constrCur
                    Case 6

                        dtbl = tblCorporateAddresses
                        var8 = frmH.cbxSubmittedTo.Text
                        If IsDBNull(var8) Then
                            varReplace = "[NA]"
                        ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                            varReplace = "[NA]"
                        Else
                            str1 = var8
                            str2 = "charNickName = '" & str1 & "'"
                            tblNick = tblCorporateNickNames
                            rowsNick = tblNick.Select(str2)
                            var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                            str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1"
                            drows = dtbl.Select(str2, "id_tbladdresslabels ASC")
                            int1 = drows.Length
                            If int1 < 1 Then
                                varReplace = "[NA]"
                            Else
                                var8 = ""
                                For Count2 = 0 To int1 - 1
                                    var8 = var8 & drows(Count2).Item("charValue") & " "
                                Next
                                'remove blank space at end
                                var8 = Trim(var8)
                                varReplace = var8
                            End If
                        End If

                        str1 = "Sponsor:"
                        var1 = var8
                    Case 7

                        str1 = NZ(frmH.cbxAnticoagulant.Text, "")
                        'determine if text should be capitalized
                        var8 = "[NA]"
                        If Len(str1) = 0 Then
                        Else
                            int1 = Len(str1)
                            int2 = 0
                            For Count2 = 1 To int1
                                var1 = CStr(Mid(str1, Count2, 1))
                                var2 = Asc(var1)
                                If var2 > 64 And var2 < 91 Then
                                    int2 = int2 + 1
                                End If
                            Next
                            var8 = str1
                            'If int2 > 2 Then 'probably is an acronym, leave capitalized
                            '    var8 = str1
                            'Else
                            '    var8 = UnCapit(str1, True)
                            'End If
                        End If
                        If IsDBNull(var8) Then
                            varReplace = "[NA]"
                        ElseIf Len(var8) = 0 Or StrComp(NZ(var8, "[None]"), "[None]", CompareMethod.Text) = 0 Then
                            varReplace = "[NA]"
                        Else
                            varReplace = var8
                        End If

                        str1 = "Anticoagulant:"
                        var1 = var8
                    Case 8

                        dv = frmH.dgvWatsonAnalRef.DataSource
                        int1 = FindRowDV("LLOQ Units", dv)
                        var1 = NZ(dv(int1).Item(1).ToString, "[NA]")
                        str1 = "Concentration Units:"

                    Case 9
                        str1 = "StudyDoc Study ID:"
                        var1 = id_tblStudies
                    Case 10
                        str1 = "Data Sig Figs:"
                        var1 = LSigFig
                    Case 11
                        int1 = frmH.dgvwStudy.CurrentCell.RowIndex
                        str1 = "Watson Project:"
                        var1 = frmH.dgvwStudy("ProjectIDText", int1).Value
                    Case 12
                        int1 = frmH.dgvwStudy.CurrentCell.RowIndex
                        str1 = "Watson Study"
                        var1 = frmH.dgvwStudy("StudyName", int1).Value
                End Select

                If boolDo Then
                    tbl.Cell(Count1, 1).Select()
                    .Selection.Text = str1
                    tbl.Cell(Count1, 2).Select()
                    .Selection.Text = var1
                End If

            Next

            'wd.Visible = True


            'start entering cmpd info
            tbl = wd.ActiveDocument.Tables(2)
            'wd.Visible = True
            tblRows = tbl.Rows.Count
            intRows = tblAnalytesHome.Rows.Count

            If intRows > tblRows Then
                tbl.Cell(tblRows, 1).Select()
                wd.Selection.InsertRowsBelow(intRows - tblRows)
            End If

            '20160202 LEE: don't add Group column
            tblCols = tblAnalytesHome.Columns.Count - 1
            intCols = tbl.Columns.Count
            int1 = tblCols - intCols
            If tblCols > intCols Then
                For Count1 = 1 To int1
                    tbl.Cell(1, intCols).Select()
                    wd.Selection.InsertColumnsRight()
                    intCols = tbl.Columns.Count
                Next
            End If

            'enter analyte information
            For Count1 = 1 To tblAnalytesHome.Rows.Count
                '20160202 LEE: don't add Group column
                For Count2 = 1 To tblAnalytesHome.Columns.Count - 1
                    var1 = tblAnalytesHome.Rows(Count1 - 1).Item(Count2 - 1)
                    'tbl.Cell(Count1, Count2).Select()
                    var2 = NZ(var1, "")
                    'var3 = Replace(var2, ChrW(173), ChrW(45), 1, -1, CompareMethod.Text)
                    var3 = var2 'Replace isn't working correctly for some reason: Replace(var2, ChrW(173), ChrW(45), 1, -1, CompareMethod.Text)

                    tbl.Cell(Count1, Count2).Range.Text = var3

                    '.Selection.Text = var3
                    var2 = var1 'debug
                Next
            Next

        End With

        wd.ActiveDocument.Save()

        boolHit = False

        For Count1 = Len(strP) To 1 Step -1
            str1 = Mid(strP, Count1, 1)
            If StrComp(str1, "\", CompareMethod.Text) = 0 Then
                intHit = Count1
                boolHit = True
                Exit For
            End If
        Next

        If boolHit Then
            strName = Mid(strP, intHit + 1, Len(strP))
        Else
            strName = strP
        End If

end2:

        If boolFound Then

            Try
                wd.Application.Run(MacroName:=strR)
                wd.ActiveDocument.Save()
            Catch ex As Exception
                MsgBox(Err.Description)
                Try
                    wd.ActiveDocument.Save()
                Catch ex1 As Exception

                End Try
            End Try
        Else
            'strR = "DoGuWu"
            'Try
            '    wd.Application.Run(MacroName:=strR)
            '    wd.ActiveDocument.Save()
            'Catch ex As Exception
            '    strR = "DoStudyDoc"
            '    Try
            '        wd.Application.Run(MacroName:=strR)
            '        wd.ActiveDocument.Save()
            '    Catch ex1 As Exception

            '    End Try
            '    MsgBox(Err.Description)
            'End Try
        End If


end1:

    End Sub

    ' Nick Addition
    Sub SetToEditMode()

        frmH.cmdEdit.Enabled = False
        frmH.cmdEdit.BackColor = System.Drawing.Color.Gray
        frmH.cmdSave.Enabled = True
        frmH.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        frmH.cmdCancel.Enabled = True
        frmH.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        frmH.cmdExit.Enabled = False
        frmH.cmdExit.BackColor = System.Drawing.Color.Gray
        frmH.cmdRefresh.Enabled = False
        frmH.cmdRefresh.BackColor = System.Drawing.Color.Gainsboro

    End Sub

    Sub SetToNonEditMode()

        frmH.cmdEdit.Enabled = True
        frmH.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        frmH.cmdSave.Enabled = False
        frmH.cmdSave.BackColor = System.Drawing.Color.Gray
        frmH.cmdCancel.Enabled = False
        frmH.cmdCancel.BackColor = System.Drawing.Color.Gray
        frmH.cmdExit.Enabled = True
        frmH.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        frmH.cmdRefresh.Enabled = True
        frmH.cmdRefresh.BackColor = System.Drawing.Color.Gray

        'Clear the Text Filter Sample field
        frmH.txtFilterSamples.Clear()
    End Sub

    Sub FillFCRW()

        Dim var1, var2

        Try

            Dim dgv As DataGridView
            Dim dv As System.Data.DataView
            Dim strF As String
            Dim strS As String
            Dim strF1 As String
            Dim strS1 As String
            Dim Count1 As Short
            Dim Count2 As Short
            Dim str1 As String
            Dim str2 As String
            Dim int1 As Short
            Dim dtbl As System.Data.DataTable = tblCustomFieldCodes
            Dim dtbl2 As System.Data.DataTable = tblFieldCodes
            Dim rows() As DataRow
            Dim rows1() As DataRow

            'add unbound columns to tblCustomFieldCodes

            str1 = "CHKINCLUDE"
            str2 = "Include"
            If dtbl.Columns.Contains(str1) Then
            Else
                Dim nc As New DataColumn
                nc.ColumnName = str1
                nc.DataType = System.Type.GetType("System.Boolean")
                nc.Caption = str2
                nc.ReadOnly = False
                nc.AllowDBNull = False
                nc.DefaultValue = 0
                dtbl.Columns.Add(nc)

            End If


            str1 = "CHARFIELDCODE"
            str2 = "Field Code"
            If dtbl.Columns.Contains(str1) Then
            Else
                Dim nc As New DataColumn
                nc.ColumnName = str1
                nc.DataType = System.Type.GetType("System.String")
                nc.Caption = str2
                nc.ReadOnly = False
                dtbl.Columns.Add(nc)
            End If

            str1 = "CHARDESCRIPTION"
            str2 = "Description"
            If dtbl.Columns.Contains(str1) Then
            Else
                Dim nc As New DataColumn
                nc.ColumnName = str1
                nc.DataType = System.Type.GetType("System.String")
                nc.Caption = str2
                nc.ReadOnly = False
                dtbl.Columns.Add(nc)
            End If

            str1 = "CHAREXAMPLE"
            str2 = "Example"
            If dtbl.Columns.Contains(str1) Then
            Else
                Dim nc As New DataColumn
                nc.ColumnName = str1
                nc.DataType = System.Type.GetType("System.String")
                nc.Caption = str2
                nc.ReadOnly = False
                dtbl.Columns.Add(nc)
            End If

            'populate data
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            strS = "ID_TBLCUSTOMFIELDCODES ASC"
            rows = dtbl.Select(strF, strS) 'tblCustomFieldCodes

            'debugging
            For Count1 = 0 To dtbl2.Columns.Count - 1
                str1 = dtbl2.Columns(Count1).ColumnName
                str2 = str1
            Next


            For Count1 = 0 To rows.Length - 1
                var1 = rows(Count1).Item("ID_TBLFIELDCODES")
                strF1 = "ID_TBLFIELDCODES = " & var1
                rows1 = dtbl2.Select(strF1) 'tblFieldCodes

                If rows1.Length = 0 Then
                Else

                    rows(Count1).BeginEdit()
                    str1 = NZ(rows1(0).Item("CHARFIELDCODE"), "NA")
                    rows(Count1).Item("CHARFIELDCODE") = str1
                    str1 = NZ(rows1(0).Item("CHARDESCRIPTION"), "NA")
                    rows(Count1).Item("CHARDESCRIPTION") = str1
                    str1 = NZ(rows1(0).Item("CHAREXAMPLE"), "NA")
                    rows(Count1).Item("CHAREXAMPLE") = str1

                    rows(Count1).EndEdit()
                End If

            Next

            dgv = frmH.dgvFC

            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARFIELDCODE IS NOT NULL"
            strS = "INTORDER ASC, ID_TBLCUSTOMFIELDCODES ASC"

            'fill CHKINCLUDE
            Dim rowsI() As DataRow = tblCustomFieldCodes.Select(strF)
            For Count1 = 0 To rowsI.Length - 1
                rowsI(Count1).Item("CHKINCLUDE") = rowsI(Count1).Item("BOOLINCLUDE")
            Next

            dv = New DataView(tblCustomFieldCodes, strF, strS, DataViewRowState.CurrentRows)

            var1 = dv.Count 'debug

            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            dgv.DataSource = dv

            For Count1 = 0 To dgv.Columns.Count - 1
                dgv.Columns(Count1).Visible = False
                dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            str1 = "Include"
            str2 = "CHKINCLUDE"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            dgv.Columns(str2).ReadOnly = True
            dgv.Columns(str2).DisplayIndex = 0
            dgv.Columns(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgv.Columns(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None


            str1 = "BOOLINCLUDE"
            str2 = "BOOLINCLUDE"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = False
            dgv.Columns(str2).ReadOnly = True
            dgv.Columns(str2).DisplayIndex = 1
            dgv.Columns(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgv.Columns(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None

            str1 = "Field Code (ReadOnly)"
            str2 = "CHARFIELDCODE"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            dgv.Columns(str2).ReadOnly = True
            dgv.Columns(str2).DisplayIndex = 2
            dgv.Columns(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None

            str1 = "Value"
            str2 = "CHARVALUE"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            dgv.Columns(str2).ReadOnly = False
            dgv.Columns(str2).DisplayIndex = 3
            dgv.Columns(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            dgv.Columns(str2).DefaultCellStyle.WrapMode = DataGridViewTriState.True

            str1 = "Description (ReadOnly)"
            str2 = "CHARDESCRIPTION"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            dgv.Columns(str2).ReadOnly = True
            dgv.Columns(str2).DisplayIndex = 4
            dgv.Columns(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None

            str1 = "Example (ReadOnly)"
            str2 = "CHAREXAMPLE"
            dgv.Columns(str2).HeaderText = str1
            dgv.Columns(str2).Visible = True
            dgv.Columns(str2).ReadOnly = True
            dgv.Columns(str2).DisplayIndex = 5
            dgv.Columns(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None

            'pesky

            str2 = "CHKINCLUDE"
            dgv.Columns(str2).DisplayIndex = 0

            str2 = "BOOLINCLUDE"
            dgv.Columns(str2).DisplayIndex = 1

            str2 = "CHARFIELDCODE"
            dgv.Columns(str2).DisplayIndex = 2

            str2 = "CHARVALUE"
            dgv.Columns(str2).DisplayIndex = 3

            str2 = "CHARDESCRIPTION"
            dgv.Columns(str2).DisplayIndex = 4

            str2 = "CHAREXAMPLE"
            dgv.Columns(str2).DisplayIndex = 5

            dgv.Columns("INTORDER").Visible = True
            dgv.Columns("INTORDER").SortMode = DataGridViewColumnSortMode.NotSortable
            dgv.Columns("INTORDER").HeaderText = "Order"

            'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            'pesky
            Call ResizeFC()
            Call ResizeFC()

            Call FilterFC()

            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            dgv.AutoResizeColumns()

            dgv.RowHeadersWidth = 25

        Catch ex As Exception

            var1 = ex.Message

        End Try


    End Sub

    Sub FilterFC()

        Try
            Dim dgv As DataGridView = frmH.dgvFC
            Dim dv As DataView = dgv.DataSource
            Dim strF As String
            Dim strF1 As String

            If frmH.rbAll.Checked Then
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARFIELDCODE IS NOT NULL AND BOOLINCLUDE >= -1"
            Else
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARFIELDCODE IS NOT NULL AND BOOLINCLUDE = -1"
            End If

            dv.RowFilter = strF
        Catch ex As Exception

        End Try

    End Sub

    Sub ResizeFC()
        'pesky
        Try

            Dim dgv As DataGridView

            Dim str1 As String
            Dim str2 As String
            Dim int1 As Short

            dgv = frmH.dgvFC

            str2 = "CHKINCLUDE"
            dgv.Columns(str2).DisplayIndex = 0

            str2 = "BOOLINCLUDE"
            dgv.Columns(str2).DisplayIndex = 1

            str2 = "CHARFIELDCODE"
            dgv.Columns(str2).DisplayIndex = 2

            str2 = "CHARVALUE"
            dgv.Columns(str2).DisplayIndex = 3

            str2 = "CHARDESCRIPTION"
            dgv.Columns(str2).DisplayIndex = 4

            str2 = "CHAREXAMPLE"
            dgv.Columns(str2).DisplayIndex = 5



            str2 = "CHAREXAMPLE"
            dgv.Columns(str2).DisplayIndex = 5

            str2 = "CHARDESCRIPTION"
            dgv.Columns(str2).DisplayIndex = 4

            str2 = "CHARVALUE"
            dgv.Columns(str2).DisplayIndex = 3
            dgv.Columns(str2).MinimumWidth = 250

            str2 = "CHARFIELDCODE"
            dgv.Columns(str2).DisplayIndex = 2

            str2 = "BOOLINCLUDE"
            dgv.Columns(str2).DisplayIndex = 1

            str2 = "CHKINCLUDE"
            dgv.Columns(str2).DisplayIndex = 0

            str2 = "INTORDER"
            dgv.Columns(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


            dgv.RowHeadersWidth = 25

            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AutoResizeColumns()

        Catch ex As Exception

        End Try

    End Sub

    Sub PutReadOnlyTables()

        If frmH.chkReadOnlyTables.Checked Then
            gboolReadOnlyTables = True
        Else
            gboolReadOnlyTables = False
        End If

    End Sub

    Sub GetReadOnlyTables()

        If gboolReadOnlyTables Then
            frmH.chkReadOnlyTables.Checked = True
        Else
            frmH.chkReadOnlyTables.Checked = False
        End If

    End Sub

    Public Sub AutoSizeGrid(ByVal MaxWidth As Short, ByVal dv As System.Data.DataView, ByVal dg As DataGrid, ByVal numRows As Short, ByVal numCols As Short, ByVal intTS As Short, ByVal boolHeadersOnly As Boolean)
        Dim var1, var2, var3

        'On Error GoTo 0

        'This function will autosize the grid columns
        'This function can also be used to set column properties
        'such as null text, etc.

        'set a constant to evaluate against
        'ths text file will never have more than 300 rows

        'Const MaxWidth = 200
        Const MaxRows = 300000
        Const Padding = 15 '15

        Dim g As Graphics = Graphics.FromHwnd(dg.Handle)
        Dim sf As StringFormat = New StringFormat(StringFormat.GenericTypographic)
        Dim size As SizeF
        Dim width As Single
        Dim row As Short
        Dim col As Short
        Dim asize As SizeF

        'Dim totalrows As short
        'get total rows, total columns and the average against the MaxRows constant

        'Dim tr As short = CType(dg.DataSource, DataTable).Rows.Count
        'Dim totalrows As short
        Dim countRows As Short = IIf(numRows > MaxRows, MaxRows, numRows)
        'Dim countColumns As short = CType(dg.DataSource, DataTable).Columns.Count
        'Dim countColumns As short = dt.Columns.Count
        Dim caption As String
        Dim str1 As String
        var1 = dg.Name
        'If countRows = 0 Then
        'Else
        For col = 0 To numCols - 1

            ' Check the caption's width first 
            'caption = CType(dg.DataSource, DataTable).Columns.item(col).Caption.ToString
            'caption = dt.Columns.item(col).Caption.ToString
            caption = dg.TableStyles(intTS).GridColumnStyles(col).HeaderText
            size = g.MeasureString(caption, dg.HeaderFont, MaxWidth, sf)
            width = size.Width + Padding
            ' Loop all rows to get the widest cell
            If boolHeadersOnly Then
            Else
                For row = 0 To countRows - 1
                    'var2 = NZ(dt.Rows.item(row).Item(col), "")
                    'var2 = NZ(dv.Item(row).Item(col), "")
                    var2 = NZ(dg.Item(row, col), "")
                    If IsDate(var2) Then
                        var3 = Format(var2, "MM/dd/yyyy")
                        var2 = var3
                    End If
                    'If Len(dt.Rows.item(row).Item(col)) = 0 Then
                    If Len(var2) = 0 Then
                    Else
                        'size = g.MeasureString(dg(row, col).ToString, dg.Font, MaxWidth, sf)
                        size = g.MeasureString(var2.ToString, dg.Font, MaxWidth, sf)
                        asize = g.MeasureString(var2.ToString, dg.Font, MaxWidth, sf)
                        If (size.Width + Padding > width) Then
                            width = size.Width + Padding
                        End If
                        If asize.Width > size.Width Then
                            'double rowheight

                        End If
                    End If
                Next
            End If

            ' Apply width and apply the null text property
            ' for now, we will use an empty string "" for null text
            'If dg.TableStyles(0).GridColumnStyles(col).MappingName = "TE Order" Then
            dg.TableStyles(intTS).GridColumnStyles(col).Width = width ' CType(width, short)
            'dg.TableStyles(intTS).GridColumnStyles(col).NullText = ""
            'dg.TableStyles(intTS).GridColumnStyles(col).Alignment = HorizontalAlignment.Left


            'dg.TableStyles(0).GridColumnStyles(0).ReadOnly = True
            'End If
        Next
        'End If
        'remove any graphical components to free memory
        g.Dispose()

    End Sub

    Public Function FindHome(ByVal strSearch)
        Dim Count1 As Short
        Dim dt As System.Data.DataTable
        Dim rw As DataRow
        Dim str1 As String
        Dim str2 As String

        dt = frmH.dgHome.DataSource
        FindHome = "No Value"
        For Each rw In dt.Rows
            str1 = NZ(rw.Item(0), "")
            If StrComp(str1, strSearch, CompareMethod.Text) = 0 Then
                str2 = NZ(rw.Item(1), "")
                FindHome = str2
                Exit For
            End If
        Next
        dt = Nothing

    End Function


    Sub SaveCompanyAnalRefTable()

        Dim i As Short
        Dim drow As DataRow
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str2a As String
        Dim str3 As String
        Dim str4 As String
        Dim int1 As Short
        Dim dtbl As System.Data.DataTable
        Dim ct1 As Short
        Dim ct2 As Short
        Dim tblSource As System.Data.DataTable
        Dim drows() As DataRow
        Dim strF As String
        Dim var1, var2, var3
        Dim dcol As DataColumn
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim maxID, maxID1
        Dim maxIDI, maxID1I
        Dim tblR As System.Data.DataTable
        Dim rowsR() As DataRow
        Dim ctR As Short
        Dim boolHook As Boolean
        Dim dv As System.Data.DataView
        Dim intIDRow As Long
        Dim dvW As System.Data.DataView
        Dim strAnal As String
        Dim tblI As System.Data.DataTable
        Dim rowsI() As DataRow
        Dim rowsM() As DataRow
        Dim intRep As Short
        Dim varA, varB
        Dim strS As String
        Dim dgv As DataGridView
        Dim strColName As String

        tblI = tblIncludedRows
        dgv = frmH.dgvCompanyAnalRef
        'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARTABLENAME = 'tblAnalRefStandards'"
        'strS = "INTORDER ASC"
        'rowsI = tblI.Select(strF, strS)

        dv = frmH.dgvCompanyAnalRef.DataSource
        dvW = frmH.dgvWatsonAnalRef.DataSource
        Dim tblW As System.Data.DataTable = dvW.ToTable

        boolHook = False
        If Len(AnalRefHook) = 0 Then
        Else
            boolHook = True
        End If
        boolHook = False 'keep boolhook false for now 20060828

        tblR = tblDataTableRowTitles
        str1 = "CHARDATATABLENAME = 'tblCompanyAnalRefTable'"
        strS = "INTORDER ASC"
        rowsR = tblR.Select(str1, strS)
        ctR = rowsR.Length

        'get max id
        maxID = GetMaxID("tblAnalRefStandards", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
        maxIDI = GetMaxID("tblIncludedRows", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid

        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If
        'tbl = tblMaxID
        'str1 = "charTable = 'tblAnalRefStandards'"
        'rows = tbl.Select(str1)
        'maxID = rows(0).Item("NUMMAXID")
        maxID1 = maxID

        ''for tblI
        'str1 = "charTable = 'tblIncludedRows'"
        'rowsM = tbl.Select(str1)
        'maxIDI = rowsM(0).Item("NUMMAXID")
        maxID1I = maxIDI

        'ct1 = ctAnalytes + ctAnalytes_IS
        dtbl = tblCompanyAnalRefTable
        ct1 = dtbl.Columns.Count ' 
        tblSource = tblAnalRefStandards

        'intIDRow = FindRowDVByCol("ID", dv, "Item")
        intIDRow = FindRow("Id", dtbl, "Item")
        varA = dtbl.Rows.Item(intIDRow).Item("Item")
        varB = dv(intIDRow).Item("Item")

        intRep = 0
        For Count2 = 0 To ct1 - 1
            'var1 = arrAnalytes(1, Count2)
            strColName = dtbl.Columns.Item(Count2).ColumnName
            If StrComp(strColName, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'IGNORE
            ElseIf StrComp(strColName, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'IGNORE
            ElseIf StrComp(strColName, "Item", CompareMethod.Text) = 0 Then 'IGNORE
            Else
                intRep = intRep + 1
                var2 = NZ(dtbl.Rows.Item(intIDRow).Item(strColName), 0)
                strF = "id_tblAnalRefStandards = " & var2
                drows = tblSource.Select(strF)
                If drows.Length = 0 Then 'add new row
                    maxID = maxID + 1
                    drow = tblSource.NewRow()
                    drow.BeginEdit()
                    drow("id_tblAnalRefStandards") = maxID
                    'also add maxid to dtbl
                    dtbl.Rows.Item(intIDRow).BeginEdit()
                    dtbl.Rows.Item(intIDRow).Item(strColName) = maxID
                    dtbl.Rows.Item(intIDRow).EndEdit()
                    var2 = maxID
                    drow("id_tblStudies") = id_tblStudies
                    'drow("charAnalyteName") = var1 'frmH.dgvCompanyAnalRef.Columns.item(Count2).Name
                    drow("boolInclude") = -1
                    var1 = dtbl.Columns.Item(Count2).ColumnName
                    drow("CHARCOLUMNNAME") = var1
                Else 'edit existing row
                    drows(0).BeginEdit()
                End If

                For Count1 = 0 To ctR - 1
                    str1 = rowsR(Count1).Item("CHARROWNAME")
                    str2 = rowsR(Count1).Item("CHARTABLEREFCOLUMNNAME")
                    str2a = NZ(rowsR(Count1).Item("CHARHOOK"), "")

                    'int1 = FindRowDVByCol(str1, dv, "Item")
                    int1 = FindRow(str1, dtbl, "Item")
                    ' var2 = NZ(dtbl.Rows.Item(int1).Item(strColName), "")
                    If int1 = -1 Then
                        GoTo next1
                    End If
                    var2 = NZ(dtbl.Rows.Item(int1).Item(strColName), "")
                    str3 = Mid(str2, 1, 2) 'dt
                    str4 = Mid(str2, 1, 4) 'bool

                    If StrComp(str4, "bool", CompareMethod.Text) = 0 Then
                        If Len(var2) = 0 Then
                            var2 = 0
                        ElseIf StrComp(var2.ToString, "Yes", CompareMethod.Text) = 0 Then
                            var2 = -1 'True
                        ElseIf StrComp(var2.ToString, "No", CompareMethod.Text) = 0 Then
                            var2 = 0 'False
                        Else
                            var2 = 0
                        End If
                    ElseIf StrComp(str3, "dt", CompareMethod.Text) = 0 Then
                        If Len(var2) = 0 Then
                            var2 = DBNull.Value
                        End If
                    Else
                    End If

                    If StrComp(str1, "ID", CompareMethod.Text) = 0 Then 'skip
                        'ElseIf StrComp(str1, "Analyte Name", CompareMethod.Text) = 0 Then 'skip
                    Else
                        'If boolHook And Len(str2a) > 0 And StrComp(str1, "Company ID", CompareMethod.Text) <> 0 Then 'don't save stuff that has a CHARHOOK value in tblR
                        'Else
                        'End If
                        If drows.Length = 0 Then
                            drow(str2) = var2
                        Else
                            drows(0).Item(str2) = var2
                        End If
                    End If

                    If intRep = 1 Then 'do this only once
                        'now save BOOLINCLUDE and ID_TBLDATATABLEROWTITLES
                        var1 = NZ(dv(Count1).Item("ID_TBLDATATABLEROWTITLES"), 0)
                        strF = "ID_TBLDATATABLEROWTITLES = " & var1 & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND CHARTABLENAME = 'tblAnalRefStandards'"
                        Erase rowsI
                        rowsI = tblI.Select(strF)
                        If rowsI.Length = 0 Then 'create new row
                            Dim ri As DataRow = tblI.NewRow
                            ri.BeginEdit()
                            maxIDI = maxIDI + 1
                            ri.Item("ID_TBLINCLUDEDROWS") = maxIDI
                            ri.Item("ID_TBLDATATABLEROWTITLES") = var1
                            ri.Item("ID_TBLSTUDIES") = id_tblStudies
                            ri.Item("CHARTABLENAME") = "tblAnalRefStandards"
                            Select Case var1
                                Case Is = 1 'Company ID
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 194 'Analyte Name
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 193 'IUPAC Name
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 192 'Alias
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 2 'Lot Number
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 3 'Physical Description
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 4 'Storage Conditions
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 5 'Date Received
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 6 'Expiration/Retest Date
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 7 'Amount Received
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 8 'Manufacturer/Supplier
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 9 'Purity
                                    ri.Item("BOOLINCLUDE") = -1
                                Case Is = 10 'Water
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 195 'Is Coadministered Cmpd?
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 198 'Certificate of Analysis
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 11 'Comments
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 12 'Is Replicate?
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 191 'Is Configured in Watson?
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 196 'ID
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 197 'Analyte Parent
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 205 'Is Internal Standard?
                                    ri.Item("BOOLINCLUDE") = 0

                                Case Is = 260 'Chem structure
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 261 'Mol Formula
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 262 'Mol Wt
                                    ri.Item("BOOLINCLUDE") = 0
                                Case Is = 263 'Monoisotopic wt
                                    ri.Item("BOOLINCLUDE") = 0

                                Case Is > 0 'this if for any future additions
                                    ri.Item("BOOLINCLUDE") = 0

                            End Select
                            ri.EndEdit()
                            tblI.Rows.Add(ri)
                        Else
                            rowsI(0).BeginEdit()
                            var2 = NZ(dtbl.Rows.Item(Count1).Item("BOOLINCLUDE"), 0)
                            If var2 = True Then
                                rowsI(0).Item("BOOLINCLUDE") = -1
                            Else
                                rowsI(0).Item("BOOLINCLUDE") = 0
                            End If
                            rowsI(0).EndEdit()
                        End If
                    End If

                Next

                If drows.Length = 0 Then
                    drow.EndEdit()
                    tblSource.Rows.Add(drow)
                Else
                    drows(0).EndEdit()
                End If

            End If

next1:

        Next

        'now check to see if any records are deleted
        ct2 = tblSource.Columns.Count
        Dim drowsS() As DataRow
        Dim cols As DataColumnCollection
        cols = dtbl.Columns
        strF = "id_tblStudies = " & id_tblStudies
        drows = tblSource.Select(strF)
        ct2 = drows.Length

        'first return all analytes with this id_tblStudies
        strF = "id_tblStudies = " & id_tblStudies
        drowsS = tblSource.Select(strF)
        'now loop through these items and delete if not present in dtbl
        Dim boolAx As Boolean

        For Count1 = 0 To drowsS.Length - 1
            var1 = CLng(NZ(drowsS(Count1).Item("id_tblAnalRefStandards"), 0))
            strF = "ID = " & var1
            boolAx = False
            For Count2 = 0 To dtbl.Columns.Count - 1
                str1 = dtbl.Columns.Item(Count2).ColumnName
                If StrComp(str1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'IGNORE
                ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'IGNORE
                ElseIf StrComp(str1, "Item", CompareMethod.Text) = 0 Then 'IGNORE
                Else
                    var3 = dtbl.Rows.Item(intIDRow).Item(str1)
                    var2 = CLng(NZ(dtbl.Rows.Item(intIDRow).Item(str1), 0))
                    If var1 = var2 Then
                        boolAx = True
                        Exit For
                    End If
                End If
            Next
            If boolAx Then 'ignore
            Else 'delete
                drowsS(Count1).Delete()
            End If
        Next

        Dim dvCheck As System.Data.DataView = New DataView(tblAnalRefStandards)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblAnalRefStandards)

            If boolGuWuOracle Then
                Try 'save TBLANALREFSTANDARDS
                    ta_tblAnalRefStandards.Update(tblAnalRefStandards)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaAnal Ref Std: " & ex.Message)
                    'ds2005.TBLANALREFSTANDARDS.Merge('ds2005.TBLANALREFSTANDARDS, True)
                End Try

            ElseIf boolGuWuAccess Then
                Try 'save TBLANALREFSTANDARDS
                    ta_tblAnalRefStandardsAcc.Update(tblAnalRefStandards)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaAnal Ref Std: " & ex.Message)
                    'ds2005Acc.TBLANALREFSTANDARDS.Merge('ds2005Acc.TBLANALREFSTANDARDS, True)
                End Try

            ElseIf boolGuWuSQLServer Then
                Try 'save TBLANALREFSTANDARDS
                    ta_tblAnalRefStandardsSQLServer.Update(tblAnalRefStandards)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaAnal Ref Std: " & ex.Message)
                    'ds2005Acc.TBLANALREFSTANDARDS.Merge('ds2005Acc.TBLANALREFSTANDARDS, True)
                End Try

            End If

        End If

        'sometimes save doesn't take. Try refilling
        If boolGuWuOracle Then
            ta_tblAnalRefStandards.Fill(tblAnalRefStandards)
        ElseIf boolGuWuAccess Then
            ta_tblAnalRefStandardsAcc.Fill(tblAnalRefStandards)
        ElseIf boolGuWuSQLServer Then
            ta_tblAnalRefStandardsSQLServer.Fill(tblAnalRefStandards)
        End If

        Dim dvCheck1 As System.Data.DataView = New DataView(tblIncludedRows)
        dvCheck1.RowStateFilter = DataViewRowState.ModifiedCurrent
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblIncludedRows)

            If boolGuWuOracle Then
                Try 'save tblIncludedRows
                    ta_tblIncludedRows.Update(tblIncludedRows)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLINCLUDEDROWS.Merge('ds2005.TBLINCLUDEDROWS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try 'save tblIncludedRows
                    ta_tblIncludedRowsAcc.Update(tblIncludedRows)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLINCLUDEDROWS.Merge('ds2005Acc.TBLINCLUDEDROWS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try 'save tblIncludedRows
                    ta_tblIncludedRowsSQLServer.Update(tblIncludedRows)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLINCLUDEDROWS.Merge('ds2005Acc.TBLINCLUDEDROWS, True)
                End Try
            End If
        End If


        'update maxid
        If maxID1 = maxID Then
        Else
            Call PutMaxID("tblAnalRefStandards", maxID)

            ''str1 = "charTable = 'tblAnalRefStandards'"
            ''tbl = tblMaxID
            ''rows = tbl.Select(str1)
            'rows(0).BeginEdit()
            'rows(0).Item("NUMMAXID") = maxID
            'rows(0).EndEdit()
            'If boolGuWuOracle Then
            '    Try
            '        ta_tblMaxID.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuAccess Then
            '    Try
            '        ta_tblMaxIDAcc.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuSQLServer Then
            '    Try
            '        ta_tblMaxIDSQLServer.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'End If

        End If

        'update maxidi
        If maxID1I = maxIDI Then
        Else

            Call PutMaxID("TBLINCLUDEDROWS", maxIDI)

            ''str1 = "charTable = 'TBLINCLUDEDROWS'"
            ''tbl = tblMaxID
            ''rowsM = tbl.Select(str1)
            'rowsM(0).BeginEdit()
            'rowsM(0).Item("NUMMAXID") = maxIDI
            'rowsM(0).EndEdit()
            'If boolGuWuOracle Then
            '    Try
            '        ta_tblMaxID.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuAccess Then
            '    Try
            '        ta_tblMaxIDAcc.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuSQLServer Then
            '    Try
            '        ta_tblMaxIDSQLServer.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'End If

        End If

        're-fill tblanalref
        '20190205 LEE:
        'FillCompanyAnalRefTable may throw an error. Alturas COR-2017-02, but I can't reproduce it
        '20190227 LEE: Problem was actually in RealMethValExecute
        Try
            Call FillCompanyAnalRefTable()
        Catch ex As Exception
            var1 = var1
        End Try
        Try
            Call Update_cbxAnalytes() 'update contents of cbxAnalytes
        Catch ex As Exception
            var1 = var1
        End Try


        str1 = AnalRefHook()
        If Len(str1) > 0 Then
            're-establish comboboxes in dgv
            Select Case str1
                Case "CRLWor_AnalRefStandard"
                    Call ComboBoxCRLAnalRefFill()
                    Call PopulateFromCRLAnalRefHook()
            End Select
        End If

        '20190205 LEE:
        Try
            frmH.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
            Call SyncCols(frmH.dgvWatsonAnalRef, frmH.dgvCompanyAnalRef)
        Catch ex As Exception
            var1 = var1
        End Try
      


    End Sub

    Sub SaveMethValTab()

        Dim i As Short
        Dim drow As DataRow
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim dtbl As System.Data.DataTable
        Dim ct1 As Short
        Dim ct2 As Short
        Dim tblSource As System.Data.DataTable
        Dim drows() As DataRow
        Dim strF As String
        Dim var1, var2, var3
        Dim dcol As DataColumn
        Dim tbl As System.Data.DataTable
        Dim tblRowT As System.Data.DataTable
        Dim drowRowT() As DataRow
        Dim intRowT As Short
        Dim dv As System.Data.DataView
        Dim ctdv As Short

        'determine if meth validation
        Dim boolVal As Boolean
        Dim dgvR As DataGridView
        Dim idR As Int64
        boolVal = False
        dgvR = frmH.dgvReports
        If dgvR.Rows.Count = 0 Then
        Else
            idR = NZ(dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value, -1)
            If idR > 1 And idR < 1000 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If


        'ct1 = ctAnalytes + ctAnalytes_IS
        tbl = tblMethValExistingGuWu
        dv = frmH.dgvMethValExistingGuWu.DataSource
        ctdv = dv.Count
        dtbl = tblMethodValData
        ct1 = dtbl.Columns.Count - 1 'must account for possible addition of replicate bottles
        tblSource = tblMethodValidationData
        tblRowT = tblDataTableRowTitles
        str1 = "charDataTableName = 'tblMethodValidationData' AND charTableRefColumnName IS NOT NULL"
        drowRowT = tblRowT.Select(str1, "intOrder")
        intRowT = drowRowT.Length

        For Count2 = 1 To ct1
            var1 = dtbl.Columns.Item(Count2).ColumnName.ToString
            strF = "id_tblStudies = " & id_tblStudies & " AND intColumnNumber = " & Count2
            drows = tblSource.Select(strF)
            If drows.Length = 0 Then 'add a new row
                drow = tblSource.NewRow()
                drow("id_tblStudies") = id_tblStudies
                drow("charColumnName") = var1
                drow("intColumnNumber") = Count2
                drow("id_tblReports") = 0

            Else
            End If

            'If tbl.Rows.Count = 0 Then
            'retrieve id_tblStudies2 from tbl
            If drows.Length = 0 Then
                drow("id_tblStudies2") = 0
            Else
                drows(0).BeginEdit()
                'var1 = NZ(dv(Count2 - 1).Item("id_tblStudies"), 0)
                Try
                    var1 = NZ(dv(Count2 - 1).Item("id_tblStudies"), 0)
                Catch ex As Exception
                    var1 = id_tblStudies
                End Try
                drows(0).Item("id_tblStudies2") = var1 'tbl.Rows.item(Count2 - 1).Item("id_tblStudies")

                '20150812 Larry: CHARARCHIVEPATH has been depricated
                'If boolVal Then
                '    var1 = pArchivePath
                '    drows(0).Item("CHARARCHIVEPATH") = var1
                'Else
                '    '20150812 Larry: NO! dv doesn't account for dups
                '    var1 = NZ(dv(Count2 - 1).Item("CHARARCHIVEPATH"), 0)
                '    drows(0).Item("CHARARCHIVEPATH") = var1
                'End If

            End If
            For Count1 = 0 To intRowT - 1

                str1 = drowRowT(Count1).Item("charRowName")
                str2 = drowRowT(Count1).Item("charTableRefColumnName")

                int1 = FindRow(str1, dtbl, "Item")
                var2 = NZ(dtbl.Rows.Item(int1).Item(Count2), "")
                dcol = tblSource.Columns.Item(str2)
                var3 = dcol.DataType.ToString
                '''''''''''''''''''''''''''console.writeline(var3)

                If StrComp(var3, "System.String", CompareMethod.Text) = 0 Then
                ElseIf StrComp(var3, "System.DateTime", CompareMethod.Text) = 0 Then
                    If Len(var2) = 0 Then
                        var2 = DBNull.Value
                    End If
                ElseIf StrComp(var3, "System.Decimal", CompareMethod.Text) = 0 Then
                    If StrComp(var2, "Yes", CompareMethod.Text) = 0 Then
                        var2 = True
                    ElseIf StrComp(var2, "No", CompareMethod.Text) = 0 Then
                        var2 = False
                    Else
                    End If
                ElseIf StrComp(var3, "System.Int16", CompareMethod.Text) = 0 Then
                    If Len(var2) = 0 Then
                        var2 = DBNull.Value
                    End If
                Else
                End If
                'If StrComp(str2, "numSampleSize", CompareMethod.Text) = 0 Then
                '    If Len(var2) = 0 Or StrComp(var1, "[NA]", CompareMethod.Text) = 0 Then
                '        var2 = DBNull.Value
                '    End If
                'End If
                'If drows.Length = 0 Then
                '    drow(str2) = var2
                'Else
                '    drows(0).Item(str2) = var2
                'End If

                If StrComp(str2, "numSampleSize", CompareMethod.Text) = 0 Then
                    If Len(var2) = 0 Or StrComp(var2, "[NA]", CompareMethod.Text) = 0 Then
                        var2 = DBNull.Value
                    End If
                End If
                If drows.Length = 0 Then
                    drow(str2) = var2
                Else
                    drows(0).Item(str2) = var2
                End If

            Next
            If drows.Length = 0 Then
                tblSource.Rows.Add(drow)
            Else
                drows(0).EndEdit()
            End If
        Next

        'now check to see if any records are deleted
        ct2 = tblSource.Columns.Count
        Dim drowsS() As DataRow
        Dim cols As DataColumnCollection
        cols = dtbl.Columns
        strF = "id_tblStudies = " & id_tblStudies
        drows = tblSource.Select(strF)
        ct2 = drows.Length
        For Count1 = 0 To ct2 - 1

            var1 = NZ(drows(Count1).Item("charColumnName"), "")
            If cols.Contains(var1) Then 'ignore
            Else 'remove record from tblSource
                drows(Count1).Delete()
            End If
        Next

        Dim dvCheck As System.Data.DataView = New DataView(tblMethodValidationData)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblMethodValidationData)

            If boolGuWuOracle Then
                Try
                    ta_tblMethodValidationData.Update(tblMethodValidationData)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaMeth Val Tabl: " & ex.Message)
                    'ds2005.TBLMETHODVALIDATIONDATA.Merge('ds2005.TBLMETHODVALIDATIONDATA, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblMethodValidationDataAcc.Update(tblMethodValidationData)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaMeth Val Tabl: " & ex.Message)
                    'ds2005Acc.TBLMETHODVALIDATIONDATA.Merge('ds2005Acc.TBLMETHODVALIDATIONDATA, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblMethodValidationDataSQLServer.Update(tblMethodValidationData)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaMeth Val Tabl: " & ex.Message)
                    'ds2005Acc.TBLMETHODVALIDATIONDATA.Merge('ds2005Acc.TBLMETHODVALIDATIONDATA, True)
                End Try
            End If

        End If

        'After one save, subsequent saves don't work
        'try updating after save
        tblMethodValidationData.Clear()
        tblMethodValidationData.AcceptChanges()
        If boolGuWuOracle Then
            tblMethodValidationData.BeginLoadData()
            ta_tblMethodValidationData.Fill(tblMethodValidationData)
            tblMethodValidationData.EndLoadData()
        ElseIf boolGuWuAccess Then
            tblMethodValidationData.BeginLoadData()
            ta_tblMethodValidationDataAcc.Fill(tblMethodValidationData)
            tblMethodValidationData.EndLoadData()
        ElseIf boolGuWuSQLServer Then
            tblMethodValidationData.BeginLoadData()
            ta_tblMethodValidationDataSQLServer.Fill(tblMethodValidationData)
            tblMethodValidationData.EndLoadData()
        End If



    End Sub

    Sub FillCompanyAnalRefTable()

        Dim i As Short
        Dim dtCols As New System.Data.DataTable
        Dim drow As DataRow
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim dtbl As System.Data.DataTable
        Dim ct1 As Short
        Dim tblSource As System.Data.DataTable
        Dim drows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim var1, var2
        Dim intL As Long
        Dim dv As System.Data.DataView
        Dim intC As Short
        Dim strColName As String
        Dim intIDRow As Long
        Dim tblR As System.Data.DataTable
        Dim rowsR() As DataRow
        Dim ctR As Short

        Dim tblI As System.Data.DataTable
        Dim rowsI() As DataRow
        tblI = tblIncludedRows

        tblR = tblDataTableRowTitles
        str1 = "CHARDATATABLENAME = 'tblCompanyAnalRefTable' and BOOLINCLUDE <> 0"
        strS = "INTORDER ASC"
        rowsR = tblR.Select(str1, strS)
        ctR = rowsR.Length

        boolFromCAR = True
        'ct1 = ctAnalytes + ctAnalytes_IS
        dtbl = tblCompanyAnalRefTable
        ct1 = dtbl.Columns.Count
        tblSource = tblAnalRefStandards

        dv = frmH.dgvCompanyAnalRef.DataSource

        intIDRow = FindRow("Id", dtbl, "Item")

        'var1 = dtbl.Columns.item(Count2).Caption
        intC = 0
        For Count2 = 0 To ct1 - 1
            strColName = dtbl.Columns.Item(Count2).ColumnName
            If StrComp(strColName, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then
            ElseIf StrComp(strColName, "BOOLINCLUDE", CompareMethod.Text) = 0 Then
            ElseIf StrComp(strColName, "Item", CompareMethod.Text) = 0 Then
            Else
                intC = intC + 1
                var1 = arrAnalytes(1, intC)

                '20190206 LEE:
                strF = "id_tblStudies = " & id_tblStudies & " and CHARCOLUMNNAME = '" & CleanText(strColName) & "'"
                drows = tblSource.Select(strF)

                For Count1 = 0 To ctR - 1
                    str1 = rowsR(Count1).Item("CHARROWNAME")
                    str2 = rowsR(Count1).Item("CHARTABLEREFCOLUMNNAME")
                    intL = rowsR(Count1).Item("ID_TBLDATATABLEROWTITLES")
                    'int1 = FindRowDVByCol(intL, dv, "ID_TBLDATATABLEROWTITLES")
                    int1 = FindRow(str1, dtbl, "Item")
                    'var2 = NZ(dtbl.Rows.Item(int1).Item(strColName), "")

                    If int1 = -1 Then
                        GoTo next1
                    End If
                    var2 = NZ(dtbl.Rows.Item(int1).Item(strColName), "")
                    If drows.Length = 0 Then
                        'If Count1 = 12 Then
                        '    'drow(arr1(1, Count2)) = arr1(8, Count2)
                        '    'drow(Count2) = arr1(8, Count2)
                        '    dtbl.Rows.item(int1).Item(Count2) = arrAnalytes(8, Count2) 'boolisreplicate
                        'Else
                        '    dtbl.Rows.item(int1).Item(Count2) = ""
                        'End If

                    Else
                        var2 = NZ(drows(0).Item(str2), "")
                        'If StrComp(str1, "Is Replicate?", CompareMethod.Text) = 0 Then
                        If StrComp(Mid(str2, 1, 4), "bool", CompareMethod.Text) = 0 Then
                            If Len(var2) = 0 Then
                                var2 = "No"
                            ElseIf var2 = -1 Then
                                var2 = "Yes"
                            ElseIf var2 = 0 Then
                                var2 = "No"
                            End If
                        ElseIf StrComp(Mid(str2, 1, 2), "dt", CompareMethod.Text) = 0 Then
                            If IsDate(var2) Then
                                var2 = CDate(var2)
                                'var2 = Format(var2, "MM/dd/yyyy")
                                var2 = Format(var2, LDateFormat)
                            End If
                        End If
                        dtbl.Rows.Item(int1).BeginEdit()
                        dtbl.Rows.Item(int1).Item(Count2) = var2
                        dtbl.Rows.Item(int1).EndEdit()

                        'If drows.Length = 0 Then
                        '    drow(str2) = var2
                        'Else
                        '    drows(0).Item(str2) = var2
                        'End If

                    End If

next1:

                Next
            End If
        Next

        'now enter boolinclude and id_tbldatatablerowtitles
        For Count1 = 0 To ctR - 1
            'intL = rowsR(Count1).Item("ID_TBLDATATABLEROWTITLES")
            'strF = "CHARTABLENAME = 'tblCompanyAnalRefTable' AND ID_TBLSTUDIES = " & id_tblStudies
            'strF = strF & " AND ID_TBLDATATABLEROWTITLES = " & intL
            'Erase rowsI
            'rowsI = tblI.Select(strF)

            intL = rowsR(Count1).Item("ID_TBLDATATABLEROWTITLES")
            strF = "CHARTABLENAME = 'tblAnalRefStandards' AND ID_TBLSTUDIES = " & id_tblStudies
            strF = strF & " AND ID_TBLDATATABLEROWTITLES = " & intL
            Erase rowsI
            rowsI = tblI.Select(strF)
            ct1 = rowsI.Length
            dtbl.Rows.Item(Count1).BeginEdit()
            If ct1 = 0 Then 'enter default values
                dtbl.Rows.Item(Count1).Item("ID_TBLDATATABLEROWTITLES") = intL

                Select Case intL
                    Case Is = 1 'Company ID
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 194 'Analyte Name
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 193 'IUPAC Name
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 192 'Alias
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 2 'Lot Number
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 3 'Physical Descdtbl.Rows.item(Count1)ption
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 4 'Storage Conditions
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 5 'Date Received
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 6 'Expiration/Retest Date
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 7 'Amount Received
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 8 'Manufacturer/Supplier
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 9 'Pudtbl.Rows.item(Count1)ty
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                    Case Is = 10 'Water
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 195 'Is Coadministered Cmpd?
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 198 'Certificate of Analysis
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 11 'Comments
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 12 'Is Replicate?
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 191 'Is Configured in Watson?
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 196 'ID
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 197 'Analyte Parent
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 205 'Is Internal Standard?
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False

                    Case Is = 260 'Chemical Structure
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 261 'Mol Formula
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 262 'Mol Wt
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                    Case Is = 263 'Monoisotopic Wt
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False

                    Case Is > False 'this if for any future additions
                        dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                End Select

            Else
                int1 = rowsI(0).Item("BOOLINCLUDE")
                If int1 = -1 Then
                    dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = True
                Else
                    dtbl.Rows.Item(Count1).Item("BOOLINCLUDE") = False
                End If
                dtbl.Rows.Item(Count1).Item("ID_TBLDATATABLEROWTITLES") = rowsI(0).Item("ID_TBLDATATABLEROWTITLES")
            End If
            dtbl.Rows.Item(Count1).EndEdit()
        Next


        'check for hook
        str1 = AnalRefHook()
        If Len(str1) > 0 Then
            Select Case str1
                Case "CRLWor_AnalRefStandard"
                    Call PopulateFromCRLAnalRefHook()
            End Select
        End If

        Call HideAnalRefRows()

        'synchronize widths with WatsonAnalRef
        Dim wid1 As Single
        wid1 = 1
        Dim wid2 As Single
        wid2 = 1
        'normalize col widths

        ''''debugWriteLine("8")
        'dv = frmh.dgCompanyAnalRef.DataSource
        'Call AutoSizeGrid(100, dv, frmh.dgCompanyAnalRef, dv.Count, ct1 + 1, 0, False)

        frmH.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
        Call SyncCols(frmH.dgvWatsonAnalRef, frmH.dgvCompanyAnalRef)

        boolFromCAR = False

        'set current cell and current row#, col#
        'frmh.dgCompanyAnalRef.CurrentCell = New DataGridCell(0, 0)
        frmH.dgvCompanyAnalRef.CurrentCell = frmH.dgvCompanyAnalRef("Item", 0)
        oldCurrentRowCAR = 0
        oldCurrentColCAR = 0
        oldCurrentCellCAR = NZ(frmH.dgvCompanyAnalRef.CurrentCell.Value, "")

    End Sub


    Sub HideAnalRefRows()

        Dim int2 As Short
        Dim int1 As Short
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim dtbl As System.Data.DataTable

        'Exit Sub


        dtbl = tblCompanyAnalRefTable

        dv = frmH.dgvCompanyAnalRef.DataSource

        For Count1 = 1 To 4
            Select Case Count1
                Case 1
                    int1 = 191 '"Is Configured in Watson?"
                Case 2
                    int1 = 197 '"Analyte Parent"
                Case 3
                    int1 = 196 '"ID"
                Case 4
                    int1 = 12 '"Is Replicate?"
            End Select
            int2 = FindRowDVByCol(int1, dv, "ID_TBLDATATABLEROWTITLES")
            If int2 = -1 Then
            Else
                frmH.dgvCompanyAnalRef.Rows.Item(int2).Visible = False
            End If
        Next

    End Sub

    Sub FillWatsonAnalRefTable() 'ByRef dataReader As IDataReader)

        Dim i As Short
        Dim intNumCols As Short
        Dim dtCols As New System.Data.DataTable
        Dim drow As DataRow
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim var1
        Dim dv As System.Data.DataView
        Dim dtbl As System.Data.DataTable
        Dim ct1 As Short
        Dim boolD As Boolean

        ct1 = ctAnalytes + ctAnalytes_IS
        dtbl = tblWatsonAnalRefTable
        'ct1 = dtbl.Columns.Count

        '
        Try
            For Count1 = 1 To 22
                Select Case Count1 '
                    Case 1
                        str1 = "Is Internal Standard?"
                    Case 2
                        str1 = "Use Internal Standard?"
                    Case 3
                        str1 = "Internal Standard"
                    Case 4
                        str1 = "LLOQ"
                    Case 5
                        str1 = "LLOQ Units"
                    Case 6
                        str1 = "ULOQ"
                    Case 7
                        str1 = "ULOQ Units"
                    Case 8
                        str1 = "Calibration Levels"
                    Case 9
                        str1 = "Regression"
                    Case 10
                        str1 = "Weighting"
                    Case 11
                        str1 = "Minimum r^2"
                    Case 12
                        str1 = "Analyte Mean Accuracy Min"
                    Case 13
                        str1 = "Analyte Mean Accuracy Max"
                    Case 14
                        str1 = "Analyte Precision Min"
                    Case 15
                        str1 = "Analyte Precision Max"
                    Case 16
                        str1 = "QC Mean Accuracy Min"
                    Case 17
                        str1 = "QC Mean Accuracy Max"
                    Case 18
                        str1 = "QC Precision Min"
                    Case 19
                        str1 = "QC Precision Max"
                    Case 20
                        str1 = "# of QC Replicates"
                    Case 21
                        str1 = "# of QC Levels"
                    Case 22
                        str1 = "# of Dilution QC Replicates"

                End Select

                int2 = 0
                For Count2 = 1 To ct1
                    'check if analyte is Watson Duplicate

                    int2 = int2 + 1
                    dtbl.Columns.Item(int2).ReadOnly = False

                    'int1 = dtbl.Columns.Count'for debugging


                    str2 = ""
                    int1 = FindRow(str1, tblWatsonAnalRefTable, "Item")
                    drow = dtbl.Rows.Item(int1)
                    Select Case str1
                        Case "Is Internal Standard?"
                            str2 = arrAnalytes(9, Count2)
                        Case "Use Internal Standard?"
                            str2 = arrAnalytes(10, Count2)
                        Case "Internal Standard"
                            str2 = arrAnalytes(11, Count2)
                        Case "LLOQ"
                            str2 = NZ(arrAnalytes(4, Count2), "")
                        Case "LLOQ Units"
                            str2 = NZ(arrAnalytes(6, Count2), "")
                        Case "ULOQ"
                            str2 = NZ(arrAnalytes(5, Count2), "")
                        Case "ULOQ Units"
                            str2 = NZ(arrAnalytes(6, Count2), "")

                    End Select
                    drow.BeginEdit()
                    drow(int2) = str2
                    drow.EndEdit()
                    'dtbl.Columns.item(Count2).ReadOnly = True
                    'don't set col read-only because entire grid is read-only
                Next
            Next
        Catch ex As Exception
            var1 = ex.Message
        End Try


        dv = dtbl.DefaultView
        'frmh.dgWatsonAnalRef.DataSource = dv
        'frmh.dgWatsonAnalRef.Refresh()
        frmH.dgvWatsonAnalRef.DataSource = dv

        Call HideWatsonRows()


        '
    End Sub

    Sub UpdateProject()

        Dim str1 As String
        Dim str2 As String
        Dim var1, var2, var3, var4, var5
        Dim dt As Date
        Dim dtbl As System.Data.DataTable
        Dim drow As DataRow
        Dim dbPath As String
        Dim Count1 As Short
        Dim intMax As Long
        Dim dv As System.Data.DataView
        Dim intRow As Short
        Dim strF As String
        Dim rows() As DataRow

        dt = Now

        gdtSave = dt

        'clear audittrailtemp
        tblAuditTrailTemp.Clear()
        idSE = 0

        If boolNewOracle Then 'this stuff done already
        Else

            'update tblStudies

            intMax = GetMaxID("tblStudies", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
            '20190219 LEE: Don't need anymore. Used GetMaxID
            'Call PutMaxID("tblStudies", intMax)

            'If boolGuWuOracle Then
            '    ta_tblMaxID.Fill(tblMaxID)
            'ElseIf boolGuWuAccess Then
            '    ta_tblMaxIDAcc.Fill(tblMaxID)
            'ElseIf boolGuWuSQLServer Then
            '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
            'End If
            'strF = "charTable = 'tblStudies'"
            'rows = tblMaxID.Select(strF)
            'intMax = rows(0).Item("nummaxid")

            ''intMax = tblStudiesL.Compute("Max(id_tblStudies)", "id_tblStudies >= 0")
            'intMax = intMax + 1
            'rows(0).BeginEdit()
            'rows(0).Item("nummaxid") = intMax
            'rows(0).EndEdit()
            'If boolGuWuOracle Then
            '    Try
            '        ta_tblMaxID.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuAccess Then
            '    Try
            '        ta_tblMaxIDAcc.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuSQLServer Then
            '    Try
            '        ta_tblMaxIDSQLServer.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'End If

            dv = New DataView(frmH.dgvwStudy.DataSource)
            intRow = frmH.dgvwStudy.CurrentRow.Index
            var1 = dv(intRow).Item("STUDYID")
            var2 = dv(intRow).Item("PROJECTID")
            var3 = dv(intRow).Item("StudyName")

            dtbl = tblStudies
            drow = dtbl.NewRow
            drow.BeginEdit()
            For Count1 = 0 To 4
                str1 = NZ(tblStudiesL.Columns.Item(Count1).Caption, "")
                Select Case str1
                    Case "ID_TBLSTUDIES"
                        var4 = intMax
                    Case "INT_WATSONSTUDYID" 'INT_WATSONSTUDYID
                        var4 = var1
                    Case "INT_WATSONPROJECTID" 'INT_WATSONPROJECTID
                        var4 = var2
                    Case "CHARWATSONSTUDYNAME" 'CHARWATSONSTUDYNAME
                        var4 = var3
                        gConfigStudy = var3
                    Case "DTCONFIGURED" 'DTCONFIGURED
                        var4 = dt
                End Select
                drow.Item(Count1) = var4 'intWatsonID
            Next
            drow.EndEdit()
            dtbl.Rows.Add(drow)

            'do this to be recorded in audit trail
            id_tblStudies = intMax

            Call FillAuditTrailTemp(tblStudies)

            If boolGuWuOracle Then
                Try
                    ta_tblStudies.Update(tblStudies)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLSTUDIES.Merge('ds2005.TBLSTUDIES, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblStudiesAcc.Update(tblStudies)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLSTUDIES.Merge('ds2005Acc.TBLSTUDIES, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblStudiesSQLServer.Update(tblStudies)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLSTUDIES.Merge('ds2005Acc.TBLSTUDIES, True)
                End Try
            End If

            'record tblaudittrailtemp
            Call RecordAuditTrail(False, dt)

        End If


        Dim int1 As Int16
        Dim int2 As Int16
        int1 = tblStudies.Rows.Count
        int2 = tblStudiesL.Rows.Count

        tblStudiesL = tblStudies



        frmH.dgStudies.Refresh()

        'update contents of cbxMethVal
        frmH.cbxMethValExistingGuWu.Items.Clear()
        Dim intRows As Short
        intRows = tblStudiesL.Rows.Count
        'add [None]
        frmH.cbxMethValExistingGuWu.Items.Add("[NONE]")
        Dim rows1() As DataRow = tblStudiesL.Select("", "charWatsonStudyName ASC", DataViewRowState.CurrentRows)
        For Count1 = 0 To intRows - 1
            frmH.cbxMethValExistingGuWu.Items.Add(rows1(Count1).Item("charWatsonStudyName"))
        Next
        frmH.cbxMethValExistingGuWu.Sorted = False


    End Sub

    Sub ClearData()


        Dim Count1 As Short
        Dim Count2 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim dtbl As System.Data.DataTable
        Dim var1

        'Home tab

        'tblReports.Clear()
        'tblReports.AcceptChanges()

        'frmH.dgvReports.DataSource = Nothing

        '20161102 LEE: can't clear tblReports for some reason
        'try this
        Dim dgv As DataGridView = frmH.dgvReports
        Try
            For Count1 = 0 To dgv.Rows.Count - 1
                For Count2 = 0 To dgv.ColumnCount - 1
                    dgv(Count2, Count1).Value = DBNull.Value
                Next
            Next
        Catch ex As Exception
            var1 = ex.Message
        End Try
     


        ''Data tab
        'dtbl = tblCompanyData

        'tblCompanyData.Clear()
        'tblCompanyData.AcceptChanges()

        'int1 = dtbl.Columns.Count
        'int2 = dtbl.Rows.Count
        'For Count1 = 1 To int1 - 1 'columns
        '    For Count2 = 0 To int2 - 1 'rows
        '        dtbl.Rows(Count2).Item(Count1) = ""
        '    Next
        'Next



        'dtbl = tblWatsonData

        'tblWatsonData.Clear()
        'tblWatsonData.AcceptChanges()

        'int1 = dtbl.Columns.Count
        'int2 = dtbl.Rows.Count
        'For Count1 = 1 To int1 - 1 'columns
        '    For Count2 = 0 To int2 - 1 'rows
        '        dtbl.Rows(Count2).Item(Count1) = ""
        '    Next
        'Next

        'frmH.dgvDataCompany.DataSource = Nothing
        'frmH.dgvStudyConfig.DataSource = Nothing
        'frmH.dgDataWatson.DataSource = Nothing


        frmH.txtSubmittedBy.Text = ""
        frmH.txtSubmittedTo.Text = ""
        frmH.txtInSupportOf.Text = ""

        Try
            frmH.cbxAnticoagulant.SelectedIndex = -1

        Catch ex As Exception

        End Try

        Try
            frmH.cbxAssayTechnique.SelectedIndex = -1

        Catch ex As Exception

        End Try

        Try
            frmH.cbxAssayTechniqueAcronym.SelectedIndex = -1

        Catch ex As Exception

        End Try

        Try
            frmH.cbxSubmittedBy.SelectedIndex = -1

        Catch ex As Exception

        End Try

        Try
            frmH.cbxInSupportOf.SelectedIndex = -1

        Catch ex As Exception

        End Try

        Try
            frmH.cbxSubmittedTo.SelectedIndex = -1

        Catch ex As Exception

        End Try

        frmH.lblReportTitle.Text = ""

        'Anal Run Summary Tab
        tblAnalRunSum.Clear()
        tblAnalRunSum.AcceptChanges()
        'frmH.dgvAnalyticalRunSummary.DataSource = Nothing

        'don't clear anymore
        'messes up later code
        GoTo end1


        frmH.dgvSummaryData.DataSource = Nothing
        frmH.dgvReportTableConfiguration.DataSource = Nothing
        frmH.dgvReportTableHeaderConfig.DataSource = Nothing
        frmH.dgvReportTables.DataSource = Nothing
        frmH.dgvCompanyAnalRef.DataSource = Nothing
        frmH.dgvWatsonAnalRef.DataSource = Nothing
        frmH.dgvContributingPersonnel.DataSource = Nothing
        frmH.dgvMethodValData.DataSource = Nothing
        frmH.dgQATable.DataSource = Nothing
        frmH.dgvSampleReceipt.DataSource = Nothing
        frmH.dgvSampleReceiptWatson.DataSource = Nothing




        'Summary table' WHY AM I CLEARING THIS TABLE???
        'frmH.dgvSummaryData.DataSource = Nothing
        'tblSummaryData.Clear()

end1:



    End Sub

    Sub Set_idtblReports()
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim intRow As Short

        dgv = frmH.dgvReports

        If dgv.RowCount = 0 Then
            intRow = -1
        ElseIf dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        If intRow = -1 Then

        Else

            Try
                id_tblReports = dgv("ID_TBLREPORTS", intRow).Value
            Catch ex As Exception
                id_tblReports = -1
            End Try
        End If

    End Sub

    Sub Configure_dgvwStudy(ByVal boolW As Boolean, ByVal con As ADODB.Connection, ByVal boolansi As Boolean)

        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolRO As Boolean
        Dim rs As New ADODB.Recordset
        Dim int1 As Int64
        Dim intCol As Short
        Dim strF As String
        Dim strS As String
        Dim var1

        dgv = frmH.dgvwStudy
        'dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        'dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AllowUserToResizeColumns = False
        dgv.AllowUserToResizeRows = False
        dgv.RowHeadersWidth = 25
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        ''debug
        'str1 = "SELECT PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME, STUDY.STUDYNUMBER, CONFIGSPECIES.SPECIES, STUDY.STUDYTITLE, PROJECT.PROJECTID, STUDY.STUDYID, STUDY.SPECIESID "
        'str2 = "FROM (PROJECT INNER JOIN STUDY ON PROJECT.PROJECTID = STUDY.PROJECTID) LEFT JOIN CONFIGSPECIES ON STUDY.SPECIESID = CONFIGSPECIES.SPECIESID "
        'str3 = "ORDER BY PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME;"
        'str4 = ""
        'strSQL = str1 & str2 & str3
        ''Console.WriteLine("Configure_dgvwStudy:  " & strSQL)

        '****

        If boolAccess Then
            'str1 = "SELECT PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME, STUDY.STUDYNUMBER, CONFIGSPECIES.SPECIES, STUDY.STUDYTITLE, PROJECT.PROJECTID, STUDY.STUDYID, STUDY.SPECIESID "
            str1 = "SELECT STUDY.STUDYNAME, PROJECT.PROJECTIDTEXT, STUDY.STUDYNUMBER, CONFIGSPECIES.SPECIES, STUDY.STUDYTITLE, PROJECT.PROJECTID, STUDY.STUDYID, STUDY.SPECIESID "
            str2 = "FROM (PROJECT INNER JOIN STUDY ON PROJECT.PROJECTID = STUDY.PROJECTID) LEFT JOIN CONFIGSPECIES ON STUDY.SPECIESID = CONFIGSPECIES.SPECIESID "
            str3 = "ORDER BY PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME;"
            str4 = ""

        Else
            If boolansi Then
                'str1 = "SELECT " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".STUDY.STUDYNUMBER, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".STUDY.STUDYTITLE, " & strSchema & ".PROJECT.PROJECTID, " & strSchema & ".STUDY.STUDYID, " & strSchema & ".STUDY.SPECIESID "
                'str2 = "FROM " & strSchema & ".CONFIGSPECIES INNER JOIN (" & strSchema & ".PROJECT INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".PROJECT.PROJECTID = " & strSchema & ".STUDY.PROJECTID) ON " & strSchema & ".CONFIGSPECIES.SPECIESID = " & strSchema & ".STUDY.SPECIESID "
                'str3 = "ORDER BY UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".STUDY.STUDYNAME);"
                'str4 = ""
                'str1 = "SELECT " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".STUDY.STUDYNUMBER, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".STUDY.STUDYTITLE, " & strSchema & ".PROJECT.PROJECTID, " & strSchema & ".STUDY.STUDYID, " & strSchema & ".STUDY.SPECIESID "

                'str1 = "SELECT " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNUMBER, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".STUDY.STUDYTITLE, " & strSchema & ".PROJECT.PROJECTID, " & strSchema & ".STUDY.STUDYID, " & strSchema & ".STUDY.SPECIESID "
                'str2 = "FROM " & strSchema & ".CONFIGSPECIES INNER JOIN (" & strSchema & ".PROJECT INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".PROJECT.PROJECTID = " & strSchema & ".STUDY.PROJECTID) ON " & strSchema & ".CONFIGSPECIES.SPECIESID = " & strSchema & ".STUDY.SPECIESID "
                'str3 = "ORDER BY UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".STUDY.STUDYNAME);"
                'str4 = ""

                '20170901 LEE: modified joins to return all records
                str1 = "SELECT " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNUMBER, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".STUDY.STUDYTITLE, " & strSchema & ".PROJECT.PROJECTID, " & strSchema & ".STUDY.STUDYID, " & strSchema & ".STUDY.SPECIESID  "
                str2 = "FROM (" & strSchema & ".PROJECT INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".PROJECT.PROJECTID = " & strSchema & ".STUDY.PROJECTID) LEFT JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".STUDY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID  "
                str3 = "ORDER BY UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".STUDY.STUDYNAME);"
                str4 = ""

            Else
                'NON-ANSI
                str1 = "SELECT PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME, STUDY.STUDYNUMBER, CONFIGSPECIES.SPECIES, STUDY.STUDYTITLE, PROJECT.PROJECTID, STUDY.STUDYID, STUDY.SPECIESID "
                str2 = "FROM PROJECT, STUDY, CONFIGSPECIES "
                str2 = str2 & "WHERE (PROJECT.PROJECTID = STUDY.PROJECTID) AND STUDY.SPECIESID = CONFIGSPECIES.SPECIESID(+) "
                str3 = "ORDER BY PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME;"

            End If
        End If


        strSQL = str1 & str2 & str3
        'Console.WriteLine("Configure_dgvwStudy:  " & strSQL)

        Try
            rs.CursorLocation = CursorLocationEnum.adUseClient
            'rs.Open(strSQL, wcn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.Open(strSQL, con) ' ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
            'GoTo end1
            GoTo eee2
        End Try

        ''int1 = rs.RecordCount
        int1 = rs.RecordCount 'debug

        tblwSTUDY.Clear()
        tblwSTUDY.AcceptChanges()
        tblwSTUDY.BeginLoadData()
        daDoPr.Fill(tblwSTUDY, rs)
        tblwSTUDY.EndLoadData()

        int1 = tblwSTUDY.Rows.Count 'debug

eee2:
        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            rs.Close()
        End If


        '20190124 LEE:
        'add tblReports.CHARREPORTTYPE
        If tblwSTUDY.Columns.Contains("CHARREPORTTYPE") Then
        Else
            Try
                Dim col1 As New DataColumn
                col1.ColumnName = "CHARREPORTTYPE"
                col1.Caption = "Study Type"
                col1.DataType = System.Type.GetType("System.String")
                tblwSTUDY.Columns.Add(col1)
            Catch ex As Exception
                var1 = var1
            End Try
        End If

        'now do tblProjects
        'str1 = "SELECT " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".PROJECT.PROJECTID "
        str1 = "SELECT " & strSchema & ".PROJECT.* "
        str2 = "FROM " & strSchema & ".PROJECT "
        str3 = "ORDER BY UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT);"


        strSQL = str1 & str2 & str3
        'Console.WriteLine(strSQL)

        Try
            rs.CursorLocation = CursorLocationEnum.adUseClient
            'rs.Open(strSQL, wcn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.Open(strSQL, con) ' ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rs.ActiveConnection = Nothing

            'int1 = rs.RecordCount 'debug
            'int1 = int1

            tblwPROJECTS.Clear()
            tblwPROJECTS.AcceptChanges()
            tblwPROJECTS.BeginLoadData()
            daDoPr.Fill(tblwPROJECTS, rs)
            tblwPROJECTS.EndLoadData()

        Catch ex As Exception
            MsgBox(ex.Message)
            GoTo end1
        End Try

        '****

        rs.Close()

        rs = Nothing

        'this should be done elsewhere
        'Call ConfigStudyTable(boolW, True)


end1:

    End Sub

    Sub GetStudyInfo()

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim dbPath As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim var1, var2, var3
        Dim arr(10)
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        'Dim 'frm As New 'frmprogress_01
        Dim num1 As Long
        Dim num2 As Long
        Dim dv As System.Data.DataView
        Dim c As DataColumn
        Dim tbl As System.Data.DataTable
        Dim wid1, wid2
        Dim varID
        Dim boolContinue As Boolean
        Dim intRows As Short
        Dim intCols As Short
        Dim dgv As DataGridView
        Dim boolT As Boolean
        Dim strF As String

        Cursor.Current = Cursors.WaitCursor

        Dim boolFL As Boolean
        boolFL = boolFormLoad
        'boolFormLoad = True

        ctPB = 0
        ctPBMax = 100
        arrQCReps.Clear(arrQCReps, 0, UBound(arrQCReps, 1))
        boolRSCFill = True
        boolStopRBS = True

        dgv = frmH.dgvwStudy

        'Display an hourglass
        Cursor.Current = Cursors.WaitCursor

        'If dgv.CurrentRow Is Nothing Then
        '    Exit Sub
        'End If
        If dgv.CurrentRow Is Nothing And boolNewOracle = False Then
            Exit Sub
        End If

        cn.Open(constrCur)

        ''''''''console.writeline(constrCur)

        intCols = dgv.ColumnCount
        intRows = dgv.RowCount

        If dgv.CurrentRow Is Nothing Then
            num1 = 0
        Else
            num1 = frmH.dgvwStudy.CurrentRow.Index
        End If

        boolT = boolFormLoad
        boolFormLoad = True

        'select cbxstudy item
        If frmH.cbxStudy.Items.Count = 0 Then
        Else
            frmH.cbxStudy.SelectedIndex = num1
        End If

        boolFormLoad = boolT

        'select cbxstudy item
        'frmH.cbxStudy.SelectedIndex = num1

        num2 = CLng(frmH.txtcbxMDBSelIndex.Text)

        'If num1 = num2 And boolRefresh = False Then
        '    boolRefresh = False
        '    GoTo end1
        'Else
        '    frmH.txtcbxMDBSelIndex.Text = num1
        'End If


        If num1 = num2 And boolRefresh = False And boolNewOracle = False Then
            boolRefresh = False
            GoTo end1
        Else
            frmH.txtcbxMDBSelIndex.Text = num1
        End If

        boolRefresh = False

        If dgv.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = dgv.CurrentRow.Index
        End If


        Dim tblS As DataTable
        tblS = frmH.dgvwStudy.DataSource
        '20190124 LEE:
        'The following code doesn't work if studies have been filtered
        'Do not need to use table
        'Can get information directly from dgv        
        'wStudyID = tblS.Rows(int1).Item("STUDYID")
        'wWStudyName = tblS.Rows(int1).Item("STUDYNAME")
        Try
            wStudyID = dgv("STUDYID", int1).Value
            wWStudyName = dgv("STUDYNAME", int1).Value
        Catch ex As Exception
            var1 = var1
        End Try

    
        '***Start here 8


        '***End here 8

        'check to see if there is a project configured for this Watson Study
        Dim dRows() As DataRow
        str1 = "int_WatsonStudyID = " & wStudyID
        dRows = tblStudiesL.Select(str1)
        int1 = dRows.Length
        If int1 = 0 Then
        Else
            var1 = dRows(0).Item("ID_TBLSTUDIES")
            id_tblStudies = var1
        End If

        Dim boolRefreshA As Boolean = False

        Dim intFilter As Long
        Call ClearForm()

        Cursor.Current = Cursors.WaitCursor

        If boolNewOracle Then

            'id_tblStudies = 0
            'id_tblStudies set in  frmBrowseWatson

            int1 = id_tblStudies 'debug

            'dgvwStudy source legend
            'str1 = "SELECT PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME, STUDY.STUDYNUMBER, CONFIGSPECIES.SPECIES, STUDY.STUDYTITLE, PROJECT.PROJECTID, STUDY.STUDYID, STUDY.SPECIESID "


            'dgvw row must be selected
            For Count1 = 0 To dgv.Rows.Count - 1
                var1 = dgv("STUDYID", Count1).Value
                If var1 = gWID Then
                    'select this row
                    Dim boolTTT As Boolean = boolFormLoad
                    boolFormLoad = True
                    dgv.CurrentCell = dgv.Rows(Count1).Cells("STUDYNAME")
                    dgv.Rows(Count1).Selected = True
                    boolFormLoad = boolTTT
                    int1 = Count1
                    wStudyID = tblS.Rows(int1).Item("STUDYID")
                    wWStudyName = tblS.Rows(int1).Item("STUDYNAME")
                    Exit For
                End If
            Next

            var1 = var1 'Debug

            '*****

            Dim strCaption As String
            Dim dt As Date

            Dim bool As Boolean

            '*****
            Dim tUserID As String
            Dim tUserName As String

            tUserID = gUserID
            tUserName = gUserName

            strRFC = GetDefaultRFC()
            strMOS = GetDefaultMOS()

            gATAdds = 0
            gATDeletes = 0
            gATMods = 0

            'If gboolAuditTrail And gboolESig Then

            '    Dim frm As New frmESig

            '    frm.ShowDialog()

            '    If frm.boolCancel Then
            '        frm.Dispose()
            '        GoTo end1
            '    End If

            '    gUserID = frm.tUserID
            '    gUserName = frm.tUserName

            '    frm.Dispose()

            'End If

            Dim dt1 As DateTime
            dt1 = Now

            '*****  STARTS UpdateProjectClick063

            'Display an hourglass
            Cursor.Current = Cursors.WaitCursor

            Call UpdateProject()

            're-establish drows
            str1 = "int_WatsonStudyID = " & wStudyID
            dRows = tblStudiesL.Select(str1)
            int1 = dRows.Length

            'update tblProjectUpdate
            'dt = Now
            'tbl = tblProjectUpdate

            boolRefresh = True

            boolNewOracle = False

            'tblwStudy should be updated by now
            'select appropriate row in dgvwstudy


            'Call GetStudyInfo()

            'z888888

            'frmH.lblGuWuStudyExists.Text = "This Watson Study IS configured in StudyDoc"
            'frmH.lblGuWuStudyExists.ForeColor = Color.Blue

            varID = dRows(0).Item("id_tblStudies")

            str1 = "Retrieving data..."
            str1 = "Refreshing StudyDoc tables..."
            ' Call PositionProgress()
            frmH.lblProgress.Text = str1
            frmH.lblProgress.Visible = True
            frmH.pb1.Maximum = ctPBMax
            frmH.pb1.Value = ctPB
            frmH.pb1.Visible = True
            frmH.pb1.Refresh()
            frmH.lblProgress.Refresh()
            frmH.lbxTab1.Enabled = True

            frmH.panProgress.Visible = True
            frmH.panProgress.Refresh()

            'do this here to refresh any date from other users
            If boolGuWuOracle Then
                boolRefreshA = DAsRefresh(frmH)
            ElseIf boolGuWuAccess Then
                boolRefreshA = DAsRefreshAcc(frmH)
            ElseIf boolGuWuSQLServer Then
                boolRefreshA = DAsRefreshSQLServer(frmH)
            End If

            str1 = "Retrieving data..."
            str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
            frmH.lblProgress.Text = str1
            frmH.lblProgress.Refresh()


            id_tblStudies = varID

            Call SetToNonEditMode()

            Cursor.Current = Cursors.WaitCursor

            frmH.pb1.Value = 0

            '20180130 LEE: Must do this before DoPrepare to set LDATEFORMAT
            Call FillDataTabData(False)

            boolContinue = DoPrepare(cn)

            'z888888


            Call SetDecs()
            boolRefresh = False

            Call Set_idtblReports()

            intOTables = 0

            Dim strM As String
            strM = "You will now be asked to apply a Study Template to this study..."
            'strM = strM & ChrW(10) & ChrW(10)
            'strM = strM & "(You will be given a chance to Cancel if you wish to not apply a Study Template.)"
            MsgBox(strM, MsgBoxStyle.Information, "Apply a Study Template...")

            Call ApplyTemplateMaster(True)

            '20180321: LEE
            'need to do some table specific stuff
            Try
                Call AddAnalyteColReportTableAnalytes()

            Catch ex As Exception
                'MsgBox("There was a problem executing AddAnalyteColReportTableAnalytes." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
            End Try

            'for some reason, the following isn't happening in normal code. Call it again
            Call FillMethValExistingGuWu()
            'have to update Summary Table again
            'update Summary Table
            Cursor.Current = Cursors.WaitCursor
            Try
                Call UpdateValueSummaryTable()
            Catch ex As Exception

            End Try
            Cursor.Current = Cursors.WaitCursor

            'di_tblStudies
            MsgBox("Study Configuration Complete!", MsgBoxStyle.Information, "Project Complete...")

            'end1:
            'Switch back to the users default cursor
            Cursor.Current = Cursors.Default

            gConfigStudy = ""


            '*****  ENDS UpdateProjectClick063

        Else
            If int1 < 1 Then

                id_tblStudies = 0
                'frmh.cmdUpdateProject.Text = "Configure Study"
                Call UpdateProjectClick063()

                GoTo end2

            Else

                varID = dRows(0).Item("id_tblStudies")

                str1 = "Retrieving data..."
                str1 = "Refreshing StudyDoc tables..."
                'Call PositionProgress()
                frmH.lblProgress.Text = str1
                frmH.lblProgress.Visible = True
                frmH.pb1.Maximum = ctPBMax
                frmH.pb1.Value = ctPB
                frmH.pb1.Visible = True
                frmH.pb1.Refresh()
                frmH.lblProgress.Refresh()
                frmH.lbxTab1.Enabled = True

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

                'do this here to refresh any data from other users


                If boolGuWuOracle Then
                    boolRefreshA = DAsRefresh(frmH)
                ElseIf boolGuWuAccess Then
                    boolRefreshA = DAsRefreshAcc(frmH)
                ElseIf boolGuWuSQLServer Then
                    boolRefreshA = DAsRefreshSQLServer(frmH)
                End If

                Call CorrectActive() 'updates tblWordStatements

                Call DatabaseCorrections() 'updates database stuff

                str1 = "Retrieving data..."
                str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
                frmH.lblProgress.Text = str1
                frmH.lblProgress.Refresh()

                id_tblStudies = varID

                Call SetToNonEditMode()


                Cursor.Current = Cursors.WaitCursor

                frmH.pb1.Value = 0

                '20180130 LEE: Must do this before DoPrepare to set LDATEFORMAT
                Call FillDataTabData(False)
                boolContinue = DoPrepare(cn)

                '20180321: LEE
                'need to do some table specific stuff
                Try
                    Call AddAnalyteColReportTableAnalytes()

                Catch ex As Exception
                    'MsgBox("There was a problem executing AddAnalyteColReportTableAnalytes." & ChrW(10) & ChrW(10) & ex.Message, MsgBoxStyle.Information, "Problem...")
                End Try

            End If
        End If

        If boolContinue Then
        Else
            GoTo end1
        End If

        'do StudyDoc study-specific queries
        Call DAsRefreshSpecific()

        'must refresh 


        'MsgBox(id_tblStudies)

        'fill tblCompanyAnalRefTable
        str1 = "Preparing Analytical Reference Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        'ctPB = ctPBMax - 10 'ctPB + 1
        ctPB = frmH.pb1.Maximum - 11
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()

        Cursor.Current = Cursors.WaitCursor
        Call AddColumnsAnalRefTable()
        Cursor.Current = Cursors.WaitCursor
        Call FillCompanyAnalRefTable()
        Cursor.Current = Cursors.WaitCursor
        Call ResizeDV(frmH.dgvCompanyAnalRef, False)
        Cursor.Current = Cursors.WaitCursor
        Call ResizeDV(frmH.dgvWatsonAnalRef, False)
        Cursor.Current = Cursors.WaitCursor
        Call Update_cbxAnalytes() 'update contents of cbxAnalytes
        Cursor.Current = Cursors.WaitCursor

        'dgCompanyAnalRef and dgWatsonAnalRef columns must be synchronized
        frmH.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
        Call SyncCols(frmH.dgvWatsonAnalRef, frmH.dgvCompanyAnalRef)

        'fill analytes in Reports Table
        str1 = "Preparing Reports Table Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()

        Cursor.Current = Cursors.WaitCursor
        Call FillReportTableHome()
        Cursor.Current = Cursors.WaitCursor
        Call ReportsSelection(False)
        Cursor.Current = Cursors.WaitCursor
        'Call SetReportHistory()
        Cursor.Current = Cursors.WaitCursor
        Call FillTableReports(True)
        Cursor.Current = Cursors.WaitCursor
        Call FillTableReportsAnalytes(True)
        Cursor.Current = Cursors.WaitCursor
        Call FillTableReportDataAnalytes(True)
        Cursor.Current = Cursors.WaitCursor
        Call CheckForTblProperties(-1, 0)

        'pesky
        Call OrderReportTableConfig()

        Cursor.Current = Cursors.WaitCursor
        'filter ReportTableHeader appropriately
        Call ReportTableHeaderConfig() 'reinitialize
        Call ReportTableHeaderFilter()
        Cursor.Current = Cursors.WaitCursor


        'enter data for companydata
        str1 = "Preparing Data Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()
        Cursor.Current = Cursors.WaitCursor

        'FillDataTabData must fill before doprepare
        'DoPrepare fills tblWatsonAnalRefTable and needs some sigfig constants
        boolFormLoad = False 'set to false so txtsubmitteto gets filled
        'Call FillDataTabData(False)
        Call frmH.ConfigDropDowDGVs()


        boolFormLoad = True
        Cursor.Current = Cursors.WaitCursor

        'enter data for Summary Table
        str1 = "Preparing Sammary Data Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        frmH.lblProgress.Refresh()
        Cursor.Current = Cursors.WaitCursor
        Call ConfigureSummaryTable()
        Cursor.Current = Cursors.WaitCursor

        'enter data for analrunsum
        str1 = "Preparing Analytical Summary Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()
        Cursor.Current = Cursors.WaitCursor


        Call FillAnalRunSum()


        Cursor.Current = Cursors.WaitCursor

        'enter data for contributing personnel
        str1 = "Preparing Contributing Personnel Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()
        Call CP_FillTable()
        Cursor.Current = Cursors.WaitCursor

        'configure table for ReportStatements
        str1 = "Preparing Report Body Section Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()
        Call ReportStatementsFill()
        Call ReportStatementsFillCharSection()
        Cursor.Current = Cursors.WaitCursor


        'configure Report Table Header info
        str1 = "Preparing Configure Column Headings Tab..."
        'str1 = "Preparing Report Table Header Configuration Tab..."
        str1 = "Preparing Configure Column Headings Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()
        Call ReportTableHeaderPopulateData()
        Call ReportTableHeaderConfigPopulate()
        Cursor.Current = Cursors.WaitCursor

        'configure QA Table info
        str1 = "Preparing QA Table Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()
        Call QATableInitialize()
        Cursor.Current = Cursors.WaitCursor

        'configure QA Table info
        str1 = "Preparing Sample Receipt Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()

        Call SampleReceiptChange()
        Cursor.Current = Cursors.WaitCursor


        frmH.dgvStudyConfig.Refresh()


        'frm.Hide()
        cn.Close()

        ''update Summary Table
        'Cursor.Current = Cursors.WaitCursor
        'Try
        '    Call UpdateValueSummaryTable()
        'Catch ex As Exception

        'End Try
        'Cursor.Current = Cursors.WaitCursor

        ''update Method Validation Table
        'Call GetMethodInfo()

        'configure QA Table info
        str1 = "Preparing Assigned Samples information..."
        str1 = str1 & ChrW(10) & "...If the study is large, this step may take a few moments..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()

        Cursor.Current = Cursors.WaitCursor

        're-investigate QC info

        Call FillAssignedSamplesDGV()
        Cursor.Current = Cursors.WaitCursor
        Dim boolAS As Boolean = UpdateGroupsAssignedSamples() 'this is for old code that made IntStd sample intGroup = 0 or null
        If boolAS Then
            'do this again
            Call FillAssignedSamplesDGV()
        End If
        Cursor.Current = Cursors.WaitCursor

        Call AssessQCs()
        Cursor.Current = Cursors.WaitCursor

        'update Method Validation Table
        'do this at end
        'give tblAnalysisResultsHome a chance to be prepared
        Call GetMethodInfo()
        Cursor.Current = Cursors.WaitCursor

        'configure table in methodVal tab
        str1 = "Preparing Method Validation Tab..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        str1 = str1 & ChrW(10) & "...this may take a moment..."
        frmH.lblProgress.Text = str1
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()
        Call FillMethValExistingGuWu()
        Cursor.Current = Cursors.WaitCursor

        'update Summary Table
        Cursor.Current = Cursors.WaitCursor
        Try
            Call UpdateValueSummaryTable()
        Catch ex As Exception

        End Try
        Cursor.Current = Cursors.WaitCursor

        str1 = "Preparing additional information..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        frmH.lblProgress.Refresh()

        'redo pesky items
        Call ResizeRows(frmH.dgvCompanyAnalRef)
        Cursor.Current = Cursors.WaitCursor
        Call ResizeRows(frmH.dgvWatsonAnalRef)
        Cursor.Current = Cursors.WaitCursor
        Call ReorderSRec()
        Cursor.Current = Cursors.WaitCursor
        'Call AppendixUpdateCB()
        Cursor.Current = Cursors.WaitCursor
        Call AssessSampleAssignment()
        Cursor.Current = Cursors.WaitCursor
        Call SelectedRefresh()
        Cursor.Current = Cursors.WaitCursor
        Call OrderReportsHome()
        Cursor.Current = Cursors.WaitCursor
        frmH.dgvReportTableConfiguration.AutoResizeRows()

        Call frmH.ConfigAnalRef()

        Call ConfigLockFinalReport()

        'Call AppendixConfig()

end1:

        Call SetEntireReportRButton()

        'get gboolDisplayAttachments
        Dim ddt As System.Data.DataTable
        Dim ddRows() As DataRow
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        ddt = tblReports
        ddRows = ddt.Select(strF)
        If ddRows.Length = 0 Then
            gboolDisplayAttachments = False
            gboolReadOnlyTables = False
        Else
            var1 = ddRows(0).Item("BOOLDISPLAYATTACHMENTS")
            If NZ(var1, 0) = 0 Then
                gboolDisplayAttachments = False
            Else
                gboolDisplayAttachments = True
            End If

            var1 = ddRows(0).Item("BOOLREADONLYTABLES")
            If NZ(var1, 0) = 0 Then
                gboolReadOnlyTables = False
            Else
                gboolReadOnlyTables = True
            End If

        End If

        Call GetReadOnlyTables()

        Call frmH.SizeCompanyAnalRef()

        boolFormLoad = True
        Call RealMethValExecute(True)
        boolFormLoad = boolFL
        If frmH.rbEntireReport.Checked Then
            Call frmH.ViewSections(False)
        Else
            Call frmH.ViewSections(True)
        End If

        boolRSCFill = False
        boolStopRBS = False

        'run this again because cmdOrder needs repositioning
        Call OrderSummaryTable()

        're-do tblfieldcodes because it gets reset
        str1 = "Preparing Field Code information..."
        str1 = str1 & ChrW(10) & "...If the study is large, this step may take a few moments..."
        str1 = "Preparing study " & frmH.cbxStudy.Text & ChrW(10) & str1
        frmH.lblProgress.Text = str1
        frmH.lblProgress.Refresh()
        Cursor.Current = Cursors.WaitCursor
        Call ResetFieldCodes(True)
        Cursor.Current = Cursors.WaitCursor
        'more pesky
        Call RTFilter()

        dgv = frmH.dgvMethodValData
        dgv.AllowUserToOrderColumns = False
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'If boolAccess Then
        '    frmH.lblWarning.Visible = False
        '    frmH.lblWatsonWarning.Visible = False
        'Else
        '    frmH.lblWarning.Visible = True
        '    frmH.lblWatsonWarning.Visible = True
        '    Call CheckWatsonRecords()
        'End If

        frmH.lblWarning.Visible = True
        frmH.lblWatsonWarning.Visible = True
        Cursor.Current = Cursors.WaitCursor
        Call CheckWatsonRecords()
        Cursor.Current = Cursors.WaitCursor

        Call MethodValColor()
        Cursor.Current = Cursors.WaitCursor
        frmH.cmdReportHistory.Enabled = True
        frmH.cmdShowOutstanding.Enabled = True

        'hide some rows in dgvWatsonAnalRef
        Cursor.Current = Cursors.WaitCursor
        Call HideWatsonRows()

        Cursor.Current = Cursors.WaitCursor
        Call ColorMethodValRows()

end2:

        'frm.Visible = False
        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False

        frmH.panProgress.Visible = False
        frmH.panProgress.Refresh()

        'Display default cursor
        Cursor.Current = Cursors.Default

        frmH.Refresh()

        boolFormLoad = boolFL

        Try
            If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
                rs.Close()
            End If
            rs.ActiveConnection = Nothing

        Catch ex As Exception
            var1 = ex.Message
        End Try

        Try
            rs = Nothing
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Try
            If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
                cn.Close()
            End If
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Try
            cn = Nothing
        Catch ex As Exception
            var1 = ex.Message
        End Try


    End Sub


    Sub MethodValColor()

        Call ColorMethodValRows()

    End Sub

    Sub ReDoColor(dgv As DataGridView, intRow As Short)

        dgv.Rows(intRow).DefaultCellStyle.ForeColor = Color.FromArgb(231, 86, 56)

    End Sub

    Sub RealMethValExecute(boolAll As Boolean)

        '20181111 LEE:
        'If boolAll = True, then do all items in dgv
        'If boolFromAdvT, means coming from Adv Table Config

        Dim rowD() As DataRow
        Dim rowMVD() As DataRow
        Dim rowS() As DataRow
        Dim rowR() As DataRow
        Dim strF As String
        Dim strFA As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strSQL As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim var1, var2, var3
        Dim idS As Int32
        Dim dv As System.Data.DataView
        Dim dvMVD As System.Data.DataView
        Dim intMVCol As Short
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim strCol As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim boolS As Boolean
        Dim tbl2 As System.Data.DataTable
        Dim num1 As Int16
        Dim rowAnti() As DataRow
        Dim rowW() As DataRow
        Dim wSID As Int64
        Dim dgv As DataGridView
        Dim tblAnalRef As System.Data.DataTable
        Dim rowAnalRef() As DataRow
        Dim strFAR As String
        Dim strPathArchive As String
        Dim boolArchive As Boolean
        Dim intRows As Short
        Dim intRow As Short
        Dim dtbl As System.Data.DataTable
        Dim strAnal As String
        Dim boolGo As Boolean
        Dim varID 'this is the method validation ID_TBLSTUDIES
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim boolClear As Boolean
        Dim strM As String

        If frmH.cmdEdit.Enabled Then
            Exit Sub
        End If

        If frmH.cmdEdit.Enabled = False And frmH.cmdSave.Enabled = False Then
            Exit Sub
        End If

        Dim rowsS() As DataRow
        Dim rowsD() As DataRow

        Dim boolHit As Boolean

        dtbl = tblMethodValidationData

        'ensure there is a value in study field
        dgv = frmH.dgvMethValExistingGuWu
        dv = dgv.DataSource
        intRows = dv.Count

        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        Dim var1a
        Dim intCt As Short

        Dim intRowsS As Int32
        Dim intRowsD As Int32

        Dim rowsMV() As DataRow
        Dim intRowsMV As Int16

        Dim arrTMVCN(tblMethodValidationData.Columns.Count, 100) 'tablemethodvalidation charcolumnnames
        Dim intTMVCN As Short = 0


        '20190110 LEE:
        'The introduction of Groups makes the following code a problem
        'If the validation study is multi-species or multi-calibrlevels, CHARCOLUMNNAME will be different than Sample Analysis study
        Dim strAnalDescr As String

        boolHit = False
        intCt = 0
        For Count1 = 0 To intRows - 1

            str1 = dgv(0, Count1).Value
            '20181111 LEE:
            'Frontage has 2 analyte, 2 matrix
            'CHARCOLUMNAME includes matrix
            'This won't be found in MethVal study
            'need to look for originaldescription
            'need to get originaldescription from tblstudydocanalytes
            'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ANALYTEDESCRIPTION = '" & str1 & "'"
            '20190110 LEE: Forgot to change to ORIGINALANALYTEDESCRIPTION
            'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ORIGINALANALYTEDESCRIPTION = '" & CleanText(str1) & "'"
            '20190226 LEE:
            'NO! dgv shows ANALYTEDESCRIPTION
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ANALYTEDESCRIPTION = '" & CleanText(str1) & "'"
            Dim rowsAH() As DataRow = TBLSTUDYDOCANALYTES.Select(strF)
            If rowsAH.Length = 0 Then
                strAnal = str1
                strAnalDescr = str1
            Else
                strAnal = NZ(rowsAH(0).Item("ORIGINALANALYTEDESCRIPTION"), str1)
                strAnalDescr = NZ(rowsAH(0).Item("ANALYTEDESCRIPTION"), str1)
            End If

            If boolAll Then
            Else
                If dgv.Rows(Count1).Selected Then
                Else
                    GoTo next1
                End If
            End If


            boolClear = False
            varID = NZ(dgv("ID_TBLSTUDIES", Count1).Value, DBNull.Value) 'this is the method validation study

            Erase rowsS
            If IsDBNull(varID) Then
                boolClear = True
                intRowsS = 0
            Else
                '20190110 LEE:
                'Aack! CHARCOLUMNNAME is the only referece to analyte
                'CHARCOLUMNNAME will be like Baclofen_C1
                'Use strAnalDescr, which would be same as CHARCOLUMNNAME
                'need to get possible CHARCOLUMNNAME/strAnalDescr from methval study
                strF2 = "ID_TBLSTUDIES = " & varID & " AND ORIGINALANALYTEDESCRIPTION = '" & CleanText(strAnal) & "'"
                rowsMV = TBLSTUDYDOCANALYTES.Select(strF2)
                intRowsMV = rowsMV.Length
                If intRowsMV = 0 Then
                    boolHit = False
                    boolClear = False
                    GoTo skip1
                End If
                'strF = "ID_TBLSTUDIES = " & varID & " AND CHARCOLUMNNAME = '" & strAnal & "'"
                Dim boolGoo As Boolean = False
                For Count2 = 1 To intRowsMV
                    str1 = rowsMV(Count2 - 1).Item("ANALYTEDESCRIPTION")
                    strF = "ID_TBLSTUDIES = " & varID & " AND CHARCOLUMNNAME = '" & CleanText(str1) & "'"
                    rowsS = dtbl.Select(strF) 'dtbl = tblMethodValidationData

                    intRowsS = rowsS.Length
                    If intRowsS = 0 Then
                    Else

                        intTMVCN = intTMVCN + 1
                        For Count3 = 0 To tblMethodValidationData.Columns.Count - 1
                            'debug
                            var1 = tblMethodValidationData.Columns(Count3).ColumnName
                            If StrComp(var1, "CHARINTERQCACCRNG", CompareMethod.Text) = 0 Then
                                var1 = var1
                                var2 = rowsS(0).Item(Count3)
                                var2 = var2
                            End If
                            arrTMVCN(Count3 + 1, intTMVCN) = rowsS(0).Item(Count3)
                        Next
                        boolGoo = True

                    End If
                Next
                If boolGoo = False Then
                    boolHit = False
                    boolClear = False
                    GoTo skip1
                End If

            End If

            intCt = intCt + 1
            boolHit = True

            'strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARCOLUMNNAME = '" & str1 & "'"
            '20190110 LEE
            'strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARCOLUMNNAME = '" & CleanText(strAnal) & "'"
            '20190226 LEE:
            'NO!! CHARCOLUMNNAME is strAnalDescr
            strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARCOLUMNNAME = '" & CleanText(strAnalDescr) & "'"
            rowsD = dtbl.Select(strF1) 'this is Sample Analysis study
            intRowsD = rowsD.Length
            'erase all rows
            'For Count2 = 0 To rowsD.Length - 1
            '    rowsD(Count2).Delete()
            'Next

            '20181111 LEE:
            'Do not update dgv.columnname

            If rowsD.Length = 0 Then
                Dim nrow As DataRow = dtbl.NewRow
                nrow.BeginEdit()
                For Count2 = 0 To dtbl.Columns.Count - 1
                    strCol = dtbl.Columns(Count2).ColumnName
                    Select Case strCol
                        Case "ID_TBLSTUDIES"
                            nrow(strCol) = id_tblStudies
                            boolGo = False
                        Case "ID_TBLSTUDIES2"
                            nrow(strCol) = varID
                            boolGo = False
                        Case "ID_TBLREPORTS"
                            nrow(strCol) = 0 'default for now
                            boolGo = False
                        Case Else
                            boolGo = True

                    End Select
                    If boolGo Then
                        var1 = rowsS(0).Item(strCol)
                        nrow(strCol) = rowsS(0).Item(strCol)
                    End If
                Next
                nrow("INTCOLUMNNUMBER") = Count1 + 1
                nrow.EndEdit()
                'dtbl.Rows.Add(nrow)
                Try
                    dtbl.Rows.Add(nrow)
                Catch ex As Exception
                    var1 = "a" 'debugging
                End Try

            Else

                rowsD(0).BeginEdit()
                For Count2 = 0 To dtbl.Columns.Count - 1 'dtbl = tblMethodValidationData
                    strCol = dtbl.Columns(Count2).ColumnName
                    strCol = strCol 'debug

                    Select Case strCol
                        Case "ID_TBLSTUDIES"
                            rowsD(0).Item(strCol) = id_tblStudies
                            boolGo = False
                        Case "ID_TBLSTUDIES2"
                            rowsD(0).Item(strCol) = varID
                            boolGo = False
                        Case "ID_TBLREPORTS"
                            rowsD(0).Item(strCol) = 0 'default for now
                            boolGo = False
                        Case Else
                            boolGo = True

                    End Select

                    If boolGo Then
                        If boolClear Then
                            Select Case strCol
                                Case "INTCOLUMNNUMBER"
                                Case "CHARCOLUMNNAME"
                                Case Else
                                    rowsD(0).Item(strCol) = DBNull.Value
                            End Select

                        Else
                            var1 = rowsS(0).Item(strCol)

                            var1a = rowsD(0).Item(strCol) 'debug
                            'rowsD(0).Item(strCol) = rowsS(0).Item(strCol)

                            '20181111 LEE:
                            'Do not update CHARCOLUMNNAME
                            '20180301 LEE:
                            'incorrect logic, shouldn't be in a for-next loop
                            'should be accessing intTMVCN in arrtmvcn

                            'Select Case strCol
                            '    Case "CHARCOLUMNNAME"
                            '    Case Else
                            '        var1 = arrTMVCN(Count2 + 1, intTMVCN)
                            '        If StrComp(strCol, "CHARINTERQCACCRNG", CompareMethod.Text) = 0 Then
                            '            var1 = var1
                            '            var2 = var2
                            '        End If

                            '        rowsD(0).Item(strCol) = var1 ' rowsS(0).Item(strCol)
                            '        var2 = NZ(var1, "")
                            '        If Len(var2) = 0 Then
                            '        Else
                            '            Exit For
                            '        End If
                            'End Select

                            '20190301 LEE: must only look at current item
                            'For Count3 = 1 To intTMVCN
                            For Count3 = intTMVCN To intTMVCN

                                var1a = rowsD(0).Item(strCol) 'debug
                                'rowsD(0).Item(strCol) = rowsS(0).Item(strCol)

                                '20181111 LEE:
                                'Do not update CHARCOLUMNNAME
                                Select Case strCol
                                    Case "CHARCOLUMNNAME"
                                    Case Else
                                        var1 = arrTMVCN(Count2 + 1, Count3)
                                        If StrComp(strCol, "CHARINTERQCACCRNG", CompareMethod.Text) = 0 Then
                                            var1 = var1
                                            var2 = var2
                                        End If

                                        rowsD(0).Item(strCol) = var1 ' rowsS(0).Item(strCol)
                                        var2 = NZ(var1, "")
                                        If Len(var2) = 0 Then
                                        Else
                                            Exit For
                                        End If
                                End Select

                            Next Count3


                            'var1a = rowsD(0).Item(strCol) 'debug
                            ''rowsD(0).Item(strCol) = rowsS(0).Item(strCol)

                            ''20181111 LEE:
                            ''Do not update CHARCOLUMNNAME
                            'Select Case strCol
                            '    Case "CHARCOLUMNNAME"
                            '    Case Else
                            '        var1 = arrTMVCN(Count2, Count3)
                            '        rowsD(0).Item(strCol) = var1 ' rowsS(0).Item(strCol)
                            'End Select
                        End If

                    End If

                Next
                rowsD(0).Item("INTCOLUMNNUMBER") = Count1 + 1
                Try
                    rowsD(0).EndEdit()
                Catch ex As Exception

                End Try

            End If

next1:

        Next Count1

skip1:

        If boolHit Or boolClear = True Then

            Call FillMethValExistingGuWu()
            Try
                Call UpdateValueSummaryTable()
            Catch ex As Exception

            End Try

            frmH.dgvMethodValData.AutoResizeRows()
        Else
            '20190227 LEE
            'should have recorded this on 20190205
            'If RealMethValExecute is called during ApplyTemplate, this code may give false negatives for boolHit
            'BoolFormLoad = true has been added to this call  in ApplyTemplate to ignore this item
            If boolFormLoad Then
            Else

                strM = "The Analyte Name '" & strAnal & "' in the Sample Analysis study has no Analyte Name match in the chosen Method Validation study."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            End If
        End If



    End Sub

    Sub FillReportTableHome()

        Dim str1 As String
        Dim var1
        Dim int1 As Short
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim numHt As Double

        'dim tblReportsView as system.data.dataview = tblReports.DefaultView
        str1 = "id_tblStudies = " & id_tblStudies
        'dv = New DataView(tblReports, str1, "id_tblReports", DataViewRowState.OriginalRows)
        dv = New DataView(tblReports, str1, "id_tblReports", DataViewRowState.CurrentRows)
        dv.AllowEdit = True
        dv.AllowNew = False
        dv.AllowDelete = False
        frmH.dgvReports.DataSource = dv
        'frmh.dgvReports.Columns.item("charReportTitle").Width = 100

        frmH.dgvReports.AutoResizeRows()

        If dv.Count = 0 Then
            str1 = ""
        Else
            str1 = NZ(dv.Item(0).Item("charReportTitle"), "")
        End If
        frmH.lblReportTitle.Text = str1
        gReportTitle = str1

    End Sub

    Sub SaveDataTabData()

        Dim strF As String
        Dim str1 As String
        Dim str2 As String
        Dim drows() As DataRow
        Dim var1, var1a, var2, var2a, var3
        Dim c As DataColumn
        Dim tbl As System.Data.DataTable
        Dim boolcbx As Boolean
        Dim boolCorp As Boolean
        Dim boolTable As Boolean
        Dim boolEntireR As Boolean
        Dim boolChk As Boolean
        Dim boolDt As Boolean
        Dim strFld As String
        Dim strFld1 As String
        Dim contr As Control
        Dim drows2() As DataRow
        Dim int1 As Short
        Dim intRows As Short
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim maxID
        Dim numEntireR As Short
        Dim strTN As String
        Dim boolConfig As Boolean = False
        Dim boolRoundConv As Boolean = False

        Dim dvD As System.Data.DataView = frmH.dgvStudyConfig.DataSource
        Dim dvD1 As System.Data.DataView = frmH.dgvDataCompany.DataSource

        If frmH.rbEntireReport.Checked Then
            numEntireR = -1
        Else
            numEntireR = 0
        End If

        str1 = "id_tblStudies = " & id_tblStudies
        drows = tblData.Select(str1, "id_tblData ASC")
        intRows = drows.Length
        dv = tblDataTableRowTitles.DefaultView
        'tblCompanyData.Columns.item(1).ReadOnly = False
        'drows2 = tblDataTableRowTitles.Select("tblCompanyData", "intOrder ASC")
        Dim dRows1() As DataRow
        If intRows = 0 Then
            'get maxid

            maxID = GetMaxID("tblData", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
            '20190219 LEE: Don't need anymore. Used GetMaxID
            'Call PutMaxID("tblData", maxID)

            'Dim strF As String
            'Dim rowsM() As DataRow
            'If boolGuWuOracle Then
            '    ta_tblMaxID.Fill(tblMaxID)
            'ElseIf boolGuWuAccess Then
            '    ta_tblMaxIDAcc.Fill(tblMaxID)
            'ElseIf boolGuWuSQLServer Then
            '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
            'End If
            'tbl = tblMaxID
            'strF = "charTable = 'tblData'"
            'rowsM = tbl.Select(strF)
            'maxID = rowsM(0).Item("NUMMAXID")
            'maxID = maxID + 1
            'rowsM(0).BeginEdit()
            'rowsM(0).Item("NUMMAXID") = maxID
            'rowsM(0).EndEdit()
            ''tbl.AcceptChanges()
            'If boolGuWuOracle Then
            '    Try
            '        ta_tblMaxID.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuAccess Then
            '    Try
            '        ta_tblMaxIDAcc.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'ElseIf boolGuWuSQLServer Then
            '    Try
            '        ta_tblMaxIDSQLServer.Update(tblMaxID)
            '    Catch ex As DBConcurrencyException
            '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '    End Try
            'End If


            Dim nr As DataRow = tblData.NewRow
            nr.BeginEdit()

            For Each c In tblData.Columns

                boolcbx = False
                boolCorp = False
                boolTable = False
                boolEntireR = False
                boolChk = True
                boolDt = False
                boolRoundConv = False

                Select Case c.ColumnName
                    Case "ID_TBLDATA"
                        nr.Item("ID_TBLDATA") = maxID
                    Case "ID_TBLSTUDIES"
                        nr.Item(c.ColumnName) = id_tblStudies
                    Case "ID_TBLASSAYTECHNIQUE" 'ID_TBLASSAYTECHNIQUE
                        tbl = tblDropdownBoxContent
                        strFld = "3"
                        strFld1 = "charValue"
                        contr = frmH.cbxAssayTechnique
                        boolcbx = True
                    Case "ID_TBLANTICOAGULANT" 'ID_TBLANTICOAGULANT
                        tbl = tblDropdownBoxContent
                        strFld = "1"
                        strFld1 = "charValue"
                        contr = frmH.cbxAnticoagulant
                        boolcbx = True
                    Case "CHARCORPORATESTUDYID" 'CHARCORPORATESTUDYID
                        boolTable = True

                    Case "CHARPROTOCOLNUMBER" 'CHARPROTOCOLNUMBER
                        boolTable = True
                    Case "ID_TBLVOLUMEUNITS" 'ID_TBLVOLUMEUNITS
                        tbl = tblDropdownBoxContent
                        strFld = "10"
                        strFld1 = "charValue"
                        contr = frmH.cbxSampleSizeUnits
                        boolcbx = True
                    Case "ID_TBLTEMPERATURES" 'ID_TBLTEMPERATURES
                        tbl = tblDropdownBoxContent
                        strFld = "9"
                        strFld1 = "charValue"
                        contr = frmH.cbxSampleStorageTemp
                        boolcbx = True
                    Case "ID_SUBMITTEDBY" 'ID_SUBMITTEDBY
                        contr = frmH.cbxSubmittedBy
                        boolCorp = True
                    Case "ID_SUBMITTEDTO" 'ID_SUBMITTEDTO
                        contr = frmH.cbxSubmittedTo
                        boolCorp = True
                    Case "ID_INSUPPORTOF" 'ID_INSUPPORTOF
                        contr = frmH.cbxInSupportOf
                        boolCorp = True
                    Case "CHARDATAARCHIVALLOCATION" 'CHARDATAARCHIVALLOCATION
                        boolTable = True
                    Case "CHARSPONSORSTUDYNUMBER" 'CHARSPONSORSTUDYNUMBER
                        boolTable = True
                    Case "CHARSPONSORSTUDYTITLE" 'CHARSPONSORSTUDYTITLE
                        boolTable = True
                    Case "NUMSIGFIGS" '
                        boolTable = True
                    Case "CHARDATEFORMAT" '
                        boolTable = True
                    Case "CHARTEXTDATEFORMAT" '
                        boolTable = True
                    Case "NUMDECIMALS" '
                        boolTable = True
                    Case "BOOLUSESIGFIGS" '
                        boolTable = True
                    Case "CHARTIMEZONE" 'CHARTIMEZONE
                        boolTable = True
                    Case "CHAROUTLIERMETHOD" 'CHAROUTLIERMETHOD
                        boolTable = True
                    Case "BOOLENTIREREPORT" '
                        boolEntireR = True
                    Case "BOOLUSESPECRND"
                        boolTable = True
                    Case "NUMREGRSIGFIGS"
                        boolTable = True
                    Case "NUMR2SIGFIGS"
                        boolTable = True
                    Case "CHARUNITS"
                        boolTable = True
                    Case "INTQCPERCDECPLACES"
                        boolTable = True
                    Case "BOOLQCEVENTSBORDER"
                        boolChk = True
                    Case "BOOLALLOWEXCLSAMPLES"
                        boolTable = True
                    Case "BOOLALLOWGUWUACCCRIT"
                        boolTable = True
                    Case "DTSTUDYSTARTDATE"
                        boolTable = True
                    Case "DTSTUDYENDDATE"
                        boolTable = True
                    Case "INTCOMMAFORMAT"
                        boolTable = True
                    Case "BOOLBLUEHYPERLINK"
                        boolTable = True
                    Case "BOOLREDBOLDFONT"
                        boolTable = True


                    Case "NUMSIGFIGSAREA" '
                        boolTable = True
                    Case "NUMDECIMALSAREA" '
                        boolTable = True
                    Case "BOOLUSESIGFIGSAREA" '
                        boolTable = True
                    Case "BOOLUSESPECRNDAREA"
                        boolTable = True

                    Case "NUMSIGFIGSAREARATIO" '
                        boolTable = True
                    Case "NUMDECIMALSAREARATIO" '
                        boolTable = True
                    Case "BOOLUSESIGFIGSAREARATIO" '
                        boolTable = True
                    Case "BOOLUSESPECRNDAREARATIO"
                        boolTable = True

                    Case "BOOLUSESIGFIGSREGR" '
                        boolTable = True
                    Case "NUMREGRDEC" '
                        boolTable = True

                    Case "BOOLUSEREGRSCINOT" '
                        boolTable = True

                    Case "BOOLNOMCONCPAREN"
                        boolTable = True
                    Case "CHARSTPAGE"
                        boolTable = True

                    Case "BOOLTABLEDTTIMESTAMP"
                        boolTable = True

                    Case "BOOLFOOTNOTEQCMEAN"
                        boolTable = True
                    Case "BOOLFLIPHEADER"
                        boolTable = True

                    Case "BOOLQCNA"
                        boolTable = True

                    Case "BOOLBQL"
                        boolTable = True

                    Case "CHARBQL"
                        boolTable = True


                    Case "BOOLIGNOREFC"
                        boolTable = True

                    Case "BOOLPSL"
                        boolTable = True

                    Case "CHARCAPTIONTRAILER"
                        boolTable = True

                    Case "BOOLRECSIGFIG" '
                        boolTable = True

                    Case "BOOLROUNDFIVEEVEN"
                        boolRoundConv = True

                    Case "BOOLROUNDFIVEAWAY"
                        boolRoundConv = True

                    Case "BOOLCRITFULLPREC"
                        boolRoundConv = True

                    Case "BOOLCRITROUNDED"
                        boolRoundConv = True

                    Case "BOOLMEANFULLPREC"
                        boolRoundConv = True

                    Case "BOOLMEANROUNDED"
                        boolRoundConv = True

                    Case "BOOLDIFFCOLSTATS" '
                        boolTable = True


                    Case "CHARCAPTIONFOLLOW" '
                        boolTable = True

                    Case "BOOLUSERSD" '
                        boolTable = True

                    Case "BOOLTABLELABELSECTION" '
                        boolTable = True

                    Case "NUMTABLEFONTSIZE" '
                        boolTable = True

                        '20190108 LEE:
                    Case "BOOLCALIBRTABLETITLE" '
                        boolTable = True

                End Select

                If boolTable Then

                    'var1 = NZ(drows(0).Item(c.ColumnName), "")
                    str1 = "charDataTableName = 'tblCompanyData' and charTableRefColumnName = '" & c.ColumnName & "'"
                    dv.RowFilter = str1
                    var2 = NZ(dv.Item(0).Item("charRowName"), "")

                    '****
                    int1 = FindRowDV(var2, dvD)
                    If int1 = -1 Then
                        var1 = ""
                    Else
                        var1 = frmH.dgvStudyConfig.Item(1, int1).Value
                    End If
                    '****

                    'check for LSigFigs LDateFormat
                    If StrComp(c.ColumnName, "NUMSIGFIGS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GSigFig)
                        LSigFig = var1
                    ElseIf StrComp(c.ColumnName, "CHARDATEFORMAT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GDateFormat)
                        LDateFormat = var1
                    ElseIf StrComp(c.ColumnName, "CHARTEXTDATEFORMAT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GDateFormat)
                        LTextDateFormat = var1
                        'seems that a YYYY gets in there somehow
                        LTextDateFormat = Replace(LTextDateFormat, "YYYY", "yyyy", 1, -1, CompareMethod.Binary)

                    ElseIf StrComp(c.ColumnName, "NUMDECIMALS", CompareMethod.Text) = 0 Then
                        'var1 = NZ(var1, GDec)
                        'LDec = var1
                        var1 = NZ(var1, GSigFig)
                        LDec = LSigFig ' var1
                    ElseIf StrComp(c.ColumnName, "CHARTIMEZONE", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GTimeZone)
                        LTimeZone = var1
                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGS", CompareMethod.Text) = 0 Then
                        'var1 = NZ(var1, boolGUseSigFigs)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigs = True
                        Else
                            var1 = 0
                            boolLUseSigFigs = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLUSESPECRND", CompareMethod.Text) = 0 Then

                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolWyethRounding = True
                        Else
                            var1 = 0
                            LboolWyethRounding = False
                        End If

                        '*****

                    ElseIf StrComp(c.ColumnName, "NUMSIGFIGSAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GSigFigArea)
                        LSigFigArea = var1
                    ElseIf StrComp(c.ColumnName, "NUMDECIMALSAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GDecArea)
                        LDecArea = var1
                        strAreaDec = GetAreaDecStr()
                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseSigFigsArea)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigsArea = True
                        Else
                            var1 = 0
                            boolLUseSigFigsArea = False
                        End If
                    ElseIf StrComp(c.ColumnName, "BOOLUSESPECRNDAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GboolWyethRoundingArea)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolWyethRoundingArea = True
                        Else
                            var1 = 0
                            LboolWyethRoundingArea = False
                        End If

                        '*******

                    ElseIf StrComp(c.ColumnName, "NUMSIGFIGSAREARATIO", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GSigFigAreaRatio)
                        LSigFigAreaRatio = var1

                        GDecAreaRatio = var1
                        LDecAreaRatio = var1
                        strAreaDecAreaRatio = GetAreaRatioDecStr()

                        'ElseIf StrComp(c.ColumnName, "NUMDECIMALSAREARATIO", CompareMethod.Text) = 0 Then
                        '    var1 = NZ(var1, GDecAreaRatio)
                        '    LDecAreaRatio = var1
                        '    strAreaDecAreaRatio = GetAreaRatioDecStr()
                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSAREARATIO", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseSigFigsAreaRatio)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigsAreaRatio = True
                        Else
                            var1 = 0
                            boolLUseSigFigsAreaRatio = False
                        End If
                    ElseIf StrComp(c.ColumnName, "BOOLUSESPECRNDAREARATIO", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GboolWyethRoundingAreaRatio)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolWyethRoundingAreaRatio = True
                        Else
                            var1 = 0
                            LboolWyethRoundingAreaRatio = False
                        End If

                        '*******

                        '*****

                    ElseIf StrComp(c.ColumnName, "CHARCAPTIONTRAILER", CompareMethod.Text) = 0 Then

                        'user may enter blank
                        'dgv shows this as null
                        'must convert to zero-length string
                        var1 = NZ(var1, "")
                        lcharCaptionTrailer = var1

                        '*****


                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSREGR", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseSigFigsRegr)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigsRegr = True
                        Else
                            var1 = 0
                            boolLUseSigFigsRegr = False
                        End If
                    ElseIf StrComp(c.ColumnName, "BOOLUSEREGRSCINOT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseRegrSciNot)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseRegrSciNot = True
                        Else
                            var1 = 0
                            boolLUseRegrSciNot = False
                        End If
                    ElseIf StrComp(c.ColumnName, "NUMREGRDEC", CompareMethod.Text) = 0 Then
                        'var1 = NZ(var1, GRegrDec)
                        var1 = NZ(var1, GRegrSigFigs)
                        var1 = LRegrSigFigs
                        LRegrDec = var1
                        strRegrDec = GetRegrDecStr(LRegrSigFigs)
                        '****

                    ElseIf StrComp(c.ColumnName, "NUMREGRSIGFIGS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GRegrSigFigs)
                        LRegrSigFigs = var1
                    ElseIf StrComp(c.ColumnName, "NUMR2SIGFIGS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GR2SigFigs)
                        LR2SigFigs = var1


                    ElseIf StrComp(c.ColumnName, "INTQCPERCDECPLACES", CompareMethod.Text) = 0 Then
                        intQCDec = var1
                        strQCDec = GetQCDecStr()

                    ElseIf StrComp(c.ColumnName, "BOOLALLOWEXCLSAMPLES", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, gAllowExclSamples)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LAllowExclSamples = True
                        Else
                            var1 = 0
                            LAllowExclSamples = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLALLOWGUWUACCCRIT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, gAllowGuWuAccCrit)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LAllowGuWuAccCrit = True
                        Else
                            var1 = 0
                            LAllowGuWuAccCrit = False
                        End If

                    ElseIf StrComp(c.ColumnName, "DTSTUDYSTARTDATE", CompareMethod.Text) = 0 Then
                        var1 = DBNull.Value

                    ElseIf StrComp(c.ColumnName, "DTSTUDYENDDATE", CompareMethod.Text) = 0 Then
                        var1 = DBNull.Value

                    ElseIf StrComp(c.ColumnName, "INTCOMMAFORMAT", CompareMethod.Text) = 0 Then
                        gINTCOMMAFORMAT = 0
                        var1 = 0

                    ElseIf StrComp(c.ColumnName, "BOOLBLUEHYPERLINK", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolBLUEHYPERLINK)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolBLUEHYPERLINK = True
                        Else
                            var1 = 0
                            boolBLUEHYPERLINK = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLREDBOLDFONT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolRedBoldFont)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolRedBoldFont = True
                        Else
                            var1 = 0
                            boolRedBoldFont = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLNOMCONCPAREN", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, LboolNomConcParen)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolNomConcParen = True
                        Else
                            var1 = 0
                            LboolNomConcParen = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLTABLEDTTIMESTAMP", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, LBOOLTABLEDTTIMESTAMP)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LBOOLTABLEDTTIMESTAMP = True
                        Else
                            var1 = 0
                            LBOOLTABLEDTTIMESTAMP = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLFOOTNOTEQCMEAN", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolFootNoteQCMean)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolFootNoteQCMean = True
                        Else
                            var1 = 0
                            boolFootNoteQCMean = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLFLIPHEADER", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolFlipHeader)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolFlipHeader = True
                        Else
                            var1 = 0
                            boolFlipHeader = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLQCNA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolQCNA)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolQCNA = True
                        Else
                            var1 = 0
                            boolQCNA = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLBQL", CompareMethod.Text) = 0 Then

                        GoTo skip1

                        var1 = NZ(var1, boolBQL)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolBQL = True
                        Else
                            var1 = 0
                            boolBQL = False
                        End If

                    ElseIf StrComp(c.ColumnName, "CHARBQL", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, "BQL/AQL")
                        gstrBQL = var1

                    ElseIf StrComp(c.ColumnName, "BOOLIGNOREFC", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolIgnoreFC)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolIgnoreFC = True
                        Else
                            var1 = 0
                            boolIgnoreFC = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLPSL", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolPSL)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolPSL = True
                        Else
                            var1 = 0
                            boolPSL = False
                        End If


                    ElseIf StrComp(c.ColumnName, "CHARSTPAGE", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, LcharSTPage)
                        LcharSTPage = var1

                    ElseIf StrComp(c.ColumnName, "BOOLRECSIGFIG", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLRECSIGFIG)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLRECSIGFIG = True
                        Else
                            var1 = 0
                            BOOLRECSIGFIG = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLSD2", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLSD2)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLSD2 = True
                            gSDMax = 2
                        Else
                            var1 = 0
                            BOOLSD2 = False
                            gSDMax = 3
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLDIFFCOLSTATS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLDIFFCOLSTATS)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLDIFFCOLSTATS = True
                        Else
                            var1 = 0
                            BOOLDIFFCOLSTATS = False
                        End If

                    ElseIf StrComp(c.ColumnName, "CHARCAPTIONFOLLOW", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, "Tab")
                        gstrCAPTIONFOLLOW = var1


                    ElseIf StrComp(c.ColumnName, "BOOLUSERSD", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLUSERSD)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLUSERSD = True
                        Else
                            var1 = 0
                            BOOLUSERSD = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLTABLELABELSECTION", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLTABLELABELSECTION)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLTABLELABELSECTION = True
                        Else
                            var1 = 0
                            BOOLTABLELABELSECTION = False
                        End If

                    ElseIf StrComp(c.ColumnName, "NUMTABLEFONTSIZE", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, NUMTABLEFONTSIZE)
                        NUMTABLEFONTSIZE = var1

                        '20190108 LEE:
                    ElseIf StrComp(c.ColumnName, "BOOLCALIBRTABLETITLE", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLCALIBRTABLETITLE)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLCALIBRTABLETITLE = True
                        Else
                            var1 = 0
                            BOOLCALIBRTABLETITLE = False
                        End If

                    End If

                    nr.Item(c.ColumnName) = var1

                End If 'end boolTable

                '****
                If boolRoundConv Then

                    If StrComp(c.ColumnName, "BOOLROUNDFIVEEVEN", CompareMethod.Text) = 0 Then

                        var2 = gboolRoundFiveEven

                    ElseIf StrComp(c.ColumnName, "BOOLROUNDFIVEAWAY", CompareMethod.Text) = 0 Then

                        var2 = gboolRoundFiveAway

                    ElseIf StrComp(c.ColumnName, "BOOLCRITFULLPREC", CompareMethod.Text) = 0 Then

                        var2 = gboolCritFullPrec

                    ElseIf StrComp(c.ColumnName, "BOOLCRITROUNDED", CompareMethod.Text) = 0 Then

                        var2 = gboolCritRounded

                    ElseIf StrComp(c.ColumnName, "BOOLMEANFULLPREC", CompareMethod.Text) = 0 Then

                        var2 = gboolMeanFullPrec

                    ElseIf StrComp(c.ColumnName, "BOOLMEANROUNDED", CompareMethod.Text) = 0 Then

                        var2 = gboolMeanRounded

                    End If

                    If var2 Then
                        var1 = -1
                    Else
                        var1 = 0
                    End If

                    Try
                        nr.Item(c.ColumnName) = var1
                        'check for LSigFigs LDateFormat
                    Catch ex As Exception
                        nr.Item(c.ColumnName) = 0
                    End Try

                End If
                '*****

                If boolChk Then
                    Select Case c.ColumnName
                        Case "BOOLQAEVENTBORDER"
                            nr.Item(c.ColumnName) = 0
                    End Select
                End If
                If boolcbx Then
                    'var1 = NZ(drows(0).Item(c.ColumnName), "")
                    var1 = contr.Text
                    If Len(var1) = 0 Then
                        var2 = 0 '"[None]"
                    Else
                        str2 = "id_tblDropdownBoxName = " & strFld & " AND charValue = '" & var1 & "'"
                        dRows1 = tbl.Select(str2)
                        If dRows1.Length = 0 Then
                            var2 = 0
                        Else
                            var2 = dRows1(0).Item("id_tblDropdownBoxContent")
                        End If
                    End If
                    nr.Item(c.ColumnName) = var2
                End If
                If boolCorp Then
                    'tbl = tblCorporateAddresses
                    tbl = tblCorporateNickNames
                    strFld = "charNickname"
                    'strFld1 = "id_tblCorporateAddresses"
                    strFld1 = "id_tblCorporateNickNames"
                    'var1 = NZ(drows(0).Item(c.ColumnName), "")
                    var1 = contr.Text
                    If Len(var1) = 0 Or StrComp(var1, "[None]", CompareMethod.Text) = 0 Then
                        'var2 = "[None]"
                        var2 = 0
                    Else
                        str2 = strFld & " = '" & var1 & "'"
                        dRows1 = tbl.Select(str2, "id_tblCorporateNickNames ASC")
                        If dRows1.Length = 0 Then
                            'var2 = "[None]"
                            var2 = 0
                        Else
                            var2 = dRows1(0).Item(strFld1)
                        End If
                    End If
                    'contr.Text = var2
                    nr.Item(c.ColumnName) = var2
                End If
                If boolEntireR Then
                    'make sure user has permission to modify report body section

                    Dim boolA As Short
                    boolA = BOOLREPORTBODYSECTIONS
                    If boolA = -1 Then
                        nr.Item(c.ColumnName) = numEntireR
                    End If

                End If

skip1:

            Next
            nr.EndEdit()
            tblData.Rows.Add(nr)

        Else 'edit existing row

            drows(0).BeginEdit()

            For Each c In tblData.Columns
                boolcbx = False
                boolCorp = False
                boolTable = False
                boolEntireR = False
                boolChk = False
                boolDt = False
                boolConfig = True
                boolRoundConv = False
                Select Case c.ColumnName
                    Case "ID_TBLDATA"
                        boolConfig = False
                    Case "ID_TBLSTUDIES"
                        drows(0).Item(c.ColumnName) = id_tblStudies
                        boolConfig = False
                    Case "ID_TBLASSAYTECHNIQUE" 'ID_TBLASSAYTECHNIQUE
                        tbl = tblDropdownBoxContent
                        strFld = "3"
                        strFld1 = "charValue"
                        contr = frmH.cbxAssayTechnique
                        boolcbx = True
                        boolConfig = False
                    Case "ID_TBLANTICOAGULANT" 'ID_ANTICOAGULANT
                        tbl = tblDropdownBoxContent
                        strFld = "1"
                        strFld1 = "charValue"
                        contr = frmH.cbxAnticoagulant
                        boolcbx = True
                        boolConfig = False
                    Case "CHARCORPORATESTUDYID" 'CHARCORPORATESTUDYID
                        boolTable = True
                        boolConfig = False

                    Case "CHARPROTOCOLNUMBER" 'CHARPROTOCOLNUMBER
                        boolTable = True
                        boolConfig = False
                    Case "ID_TBLVOLUMEUNITS" 'ID_TBLVOLUMEUNITS
                        tbl = tblDropdownBoxContent
                        strFld = "10"
                        strFld1 = "charValue"
                        contr = frmH.cbxSampleSizeUnits
                        boolcbx = True
                        boolConfig = False
                    Case "ID_TBLTEMPERATURES" 'ID_TBLTEMPERATURES
                        tbl = tblDropdownBoxContent
                        strFld = "9"
                        strFld1 = "charValue"
                        contr = frmH.cbxSampleStorageTemp
                        boolcbx = True
                        boolConfig = False
                    Case "ID_SUBMITTEDBY" 'ID_SUBMITTEDBY
                        contr = frmH.cbxSubmittedBy
                        boolCorp = True
                        boolConfig = False
                    Case "ID_SUBMITTEDTO" 'ID_SUBMITTEDTO
                        contr = frmH.cbxSubmittedTo
                        boolCorp = True
                        boolConfig = False
                    Case "ID_INSUPPORTOF" 'ID_INSUPPORTOF
                        contr = frmH.cbxInSupportOf
                        boolCorp = True
                        boolConfig = False
                    Case "CHARDATAARCHIVALLOCATION" 'CHARDATAARCHIVALLOCATION
                        boolTable = True
                        boolConfig = False

                    Case "CHARSPONSORSTUDYNUMBER" 'CHARSPONSORSTUDYNUMBER
                        boolTable = True
                        boolConfig = False
                    Case "CHARSPONSORSTUDYTITLE" 'CHARSPONSORSTUDYTITLE
                        boolTable = True
                        boolConfig = False

                    Case "NUMSIGFIGS" '
                        boolTable = True
                    Case "CHARDATEFORMAT" '
                        boolTable = True
                    Case "CHARTEXTDATEFORMAT" '
                        boolTable = True
                    Case "NUMDECIMALS" '
                        boolTable = True
                    Case "BOOLUSESIGFIGS" '
                        boolTable = True
                    Case "CHARTIMEZONE" 'CHARTIMEZONE
                        boolTable = True
                    Case "CHAROUTLIERMETHOD" 'CHAROUTLIERMETHOD
                        boolTable = True
                        boolConfig = False

                    Case "BOOLENTIREREPORT" 'CHAROUTLIERMETHOD
                        boolEntireR = True
                    Case "BOOLUSESPECRND"
                        boolTable = True
                    Case "NUMREGRSIGFIGS"
                        boolTable = True
                    Case "NUMR2SIGFIGS"
                        boolTable = True
                    Case "CHARUNITS"
                        boolTable = True
                    Case "INTQCPERCDECPLACES"
                        boolTable = True
                    Case "BOOLQAEVENTBORDER"
                        boolChk = True
                    Case "BOOLALLOWEXCLSAMPLES"
                        boolTable = True
                    Case "BOOLALLOWGUWUACCCRIT"
                        boolTable = True
                    Case "DTSTUDYSTARTDATE"
                        boolTable = True
                        boolConfig = False

                    Case "DTSTUDYENDDATE"
                        boolTable = True
                        boolConfig = False

                    Case "INTCOMMAFORMAT"
                        boolTable = True
                    Case "BOOLBLUEHYPERLINK"
                        boolTable = True
                    Case "BOOLREDBOLDFONT"
                        boolTable = True

                    Case "NUMSIGFIGSAREA" '
                        boolTable = True
                    Case "NUMDECIMALSAREA" '
                        boolTable = True
                    Case "BOOLUSESIGFIGSAREA" '
                        boolTable = True
                    Case "BOOLUSESPECRNDAREA"
                        boolTable = True

                    Case "NUMSIGFIGSAREARATIO" '
                        boolTable = True
                    Case "NUMDECIMALSAREARATIO" '
                        boolTable = True
                    Case "BOOLUSESIGFIGSAREARATIO" '
                        boolTable = True
                    Case "BOOLUSESPECRNDAREARATIO"
                        boolTable = True

                    Case "BOOLUSESIGFIGSREGR" '
                        boolTable = True
                    Case "NUMREGRDEC" '
                        boolTable = True
                    Case "BOOLUSEREGRSCINOT" '
                        boolTable = True

                    Case "BOOLNOMCONCPAREN"
                        boolTable = True
                    Case "CHARSTPAGE"
                        boolTable = True

                    Case "BOOLTABLEDTTIMESTAMP"
                        boolTable = True

                    Case "BOOLFOOTNOTEQCMEAN"
                        boolTable = True
                    Case "BOOLFLIPHEADER"
                        boolTable = True

                    Case "BOOLQCNA"
                        boolTable = True

                    Case "BOOLBQL"
                        boolTable = True
                        GoTo next1

                    Case "CHARBQL"
                        boolTable = True


                    Case "BOOLIGNOREFC"
                        boolTable = True

                    Case "BOOLPSL"
                        boolTable = True

                    Case "BOOLRECSIGFIG"
                        boolTable = True

                    Case "BOOLROUNDFIVEEVEN"
                        boolRoundConv = True

                    Case "BOOLROUNDFIVEAWAY"
                        boolRoundConv = True

                    Case "BOOLCRITFULLPREC"
                        boolRoundConv = True

                    Case "BOOLCRITROUNDED"
                        boolRoundConv = True

                    Case "BOOLMEANFULLPREC"
                        boolRoundConv = True

                    Case "BOOLMEANROUNDED"
                        boolRoundConv = True

                    Case "CHARCAPTIONTRAILER"
                        boolTable = True

                    Case "BOOLSD2"
                        boolTable = True

                    Case "BOOLDIFFCOLSTATS"
                        boolTable = True

                    Case "CHARCAPTIONFOLLOW"
                        boolTable = True


                    Case "BOOLUSERSD"
                        boolTable = True

                    Case "BOOLTABLELABELSECTION"
                        boolTable = True

                    Case "NUMTABLEFONTSIZE"
                        boolTable = True

                        '20190108 LEE:
                    Case "BOOLCALIBRTABLETITLE"
                        boolTable = True
                End Select

                If boolTable Then

                    var1 = NZ(drows(0).Item(c.ColumnName), "")
                    str1 = "charDataTableName = 'tblCompanyData' and charTableRefColumnName = '" & c.ColumnName & "'"
                    dv.RowFilter = str1
                    var2 = NZ(dv.Item(0).Item("charRowName"), "")
                    'int1 = FindRow(var2, tblCompanyData, "Item")

                    '****

                    If boolConfig Then
                        int1 = FindRowDV(var2, dvD)
                    Else
                        int1 = FindRowDV(var2, dvD1)
                    End If

                    If int1 = -1 Then
                        GoTo next1
                    End If

                    If boolConfig Then
                        var1 = frmH.dgvStudyConfig.Item(1, int1).Value
                    Else
                        var1 = frmH.dgvDataCompany.Item(1, int1).Value
                    End If
                    '****

                    'If Len(var2) = 0 Then
                    'var2 = ""
                    'Else
                    'End If
                    'var1 = frmH.dgvDataCompany.Item(1, int1).Value

                    If StrComp(c.ColumnName, "NUMSIGFIGS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GSigFig)
                        LSigFig = var1
                    ElseIf StrComp(c.ColumnName, "CHARDATEFORMAT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GDateFormat)
                        LDateFormat = var1
                    ElseIf StrComp(c.ColumnName, "CHARTEXTDATEFORMAT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GDateFormat)
                        LTextDateFormat = var1
                        'seems that a YYYY gets in there somehow
                        LTextDateFormat = Replace(LTextDateFormat, "YYYY", "yyyy", 1, -1, CompareMethod.Binary)

                    ElseIf StrComp(c.ColumnName, "NUMDECIMALS", CompareMethod.Text) = 0 Then
                        'var1 = NZ(var1, GDec)
                        'LDec = var1
                        var1 = NZ(var1, GSigFig)
                        LDec = LSigFig 'var1
                    ElseIf StrComp(c.ColumnName, "CHARTIMEZONE", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GTimeZone)
                        LTimeZone = var1
                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGS", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigs = True
                        Else
                            var1 = 0
                            boolLUseSigFigs = False
                        End If
                    ElseIf StrComp(c.ColumnName, "BOOLUSESPECRND", CompareMethod.Text) = 0 Then

                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolWyethRounding = True
                        Else
                            var1 = 0
                            LboolWyethRounding = False
                        End If

                        '*****

                    ElseIf StrComp(c.ColumnName, "NUMSIGFIGSAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GSigFigArea)
                        LSigFigArea = var1
                    ElseIf StrComp(c.ColumnName, "NUMDECIMALSAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GDecArea)
                        LDecArea = var1
                        strAreaDec = GetAreaDecStr()

                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseSigFigsArea)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigsArea = True
                        Else
                            var1 = 0
                            boolLUseSigFigsArea = False
                        End If
                    ElseIf StrComp(c.ColumnName, "BOOLUSESPECRNDAREA", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GboolWyethRoundingArea)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolWyethRoundingArea = True
                        Else
                            var1 = 0
                            LboolWyethRoundingArea = False
                        End If

                        '*****

                    ElseIf StrComp(c.ColumnName, "NUMSIGFIGSAREARATIO", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GSigFigAreaRatio)
                        LSigFigAreaRatio = var1

                        GDecAreaRatio = var1
                        LDecAreaRatio = var1
                        strAreaDecAreaRatio = GetAreaRatioDecStr()

                        'ElseIf StrComp(c.ColumnName, "NUMDECIMALSAREARATIO", CompareMethod.Text) = 0 Then
                        '    var1 = NZ(var1, GDecAreaRatio)
                        '    LDecAreaRatio = var1
                        '    strAreaDecAreaRatio = GetAreaRatioDecStr()
                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSAREARATIO", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseSigFigsAreaRatio)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigsAreaRatio = True
                        Else
                            var1 = 0
                            boolLUseSigFigsAreaRatio = False
                        End If
                    ElseIf StrComp(c.ColumnName, "BOOLUSESPECRNDAREARATIO", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GboolWyethRoundingAreaRatio)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolWyethRoundingAreaRatio = True
                        Else
                            var1 = 0
                            LboolWyethRoundingAreaRatio = False
                        End If

                        '*****

                        '****

                    ElseIf StrComp(c.ColumnName, "CHARCAPTIONTRAILER", CompareMethod.Text) = 0 Then

                        'user may enter blank
                        'dgv shows this as null
                        'must convert to zero-length string
                        var1 = NZ(var1, "")
                        lcharCaptionTrailer = var1

                        '****


                    ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSREGR", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseSigFigsRegr)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseSigFigsRegr = True
                        Else
                            var1 = 0
                            boolLUseSigFigsRegr = False
                        End If
                    ElseIf StrComp(c.ColumnName, "BOOLUSEREGRSCINOT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, boolGUseRegrSciNot)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolLUseRegrSciNot = True
                        Else
                            var1 = 0
                            boolLUseRegrSciNot = False
                        End If
                    ElseIf StrComp(c.ColumnName, "NUMREGRDEC", CompareMethod.Text) = 0 Then
                        'var1 = NZ(var1, GRegrDec)
                        var1 = NZ(var1, GRegrSigFigs)
                        LRegrDec = var1
                        strRegrDec = GetRegrDecStr(LRegrDec)
                        '*****

                    ElseIf StrComp(c.ColumnName, "NUMREGRSIGFIGS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GRegrSigFigs)
                        LRegrSigFigs = var1
                    ElseIf StrComp(c.ColumnName, "NUMR2SIGFIGS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, GR2SigFigs)
                        LR2SigFigs = var1
                    ElseIf StrComp(c.ColumnName, "INTQCPERCDECPLACES", CompareMethod.Text) = 0 Then
                        intQCDec = var1
                        strQCDec = GetQCDecStr()

                    ElseIf StrComp(c.ColumnName, "BOOLALLOWEXCLSAMPLES", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LAllowExclSamples = True
                        Else
                            var1 = 0
                            LAllowExclSamples = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLALLOWGUWUACCCRIT", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LAllowGuWuAccCrit = True
                        Else
                            var1 = 0
                            LAllowGuWuAccCrit = False
                        End If
                    ElseIf StrComp(c.ColumnName, "DTSTUDYSTARTDATE", CompareMethod.Text) = 0 Then
                        'var1 = DBNull.Value
                        If Len(NZ(var1, "")) = 0 Then
                            var1 = DBNull.Value
                        End If

                    ElseIf StrComp(c.ColumnName, "DTSTUDYENDDATE", CompareMethod.Text) = 0 Then
                        If Len(NZ(var1, "")) = 0 Then
                            var1 = DBNull.Value
                        End If

                    ElseIf StrComp(c.ColumnName, "INTCOMMAFORMAT", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, 0)
                        gINTCOMMAFORMAT = var1

                    ElseIf StrComp(c.ColumnName, "BOOLBLUEHYPERLINK", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolBLUEHYPERLINK = True
                        Else
                            var1 = 0
                            boolBLUEHYPERLINK = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLREDBOLDFONT", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolRedBoldFont = True
                        Else
                            var1 = 0
                            boolRedBoldFont = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLNOMCONCPAREN", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LboolNomConcParen = True
                        Else
                            var1 = 0
                            LboolNomConcParen = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLTABLEDTTIMESTAMP", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            LBOOLTABLEDTTIMESTAMP = True
                        Else
                            var1 = 0
                            LBOOLTABLEDTTIMESTAMP = False
                        End If


                    ElseIf StrComp(c.ColumnName, "BOOLFOOTNOTEQCMEAN", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolFootNoteQCMean = True
                        Else
                            var1 = 0
                            boolFootNoteQCMean = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLFLIPHEADER", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolFlipHeader = True
                        Else
                            var1 = 0
                            boolFlipHeader = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLQCNA", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolQCNA = True
                        Else
                            var1 = 0
                            boolQCNA = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLBQL", CompareMethod.Text) = 0 Then

                        GoTo next1

                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolBQL = True
                        Else
                            var1 = 0
                            boolBQL = False
                        End If

                    ElseIf StrComp(c.ColumnName, "CHARBQL", CompareMethod.Text) = 0 Then

                        gstrBQL = NZ(var1, "BQL/AQL")


                    ElseIf StrComp(c.ColumnName, "BOOLIGNOREFC", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolIgnoreFC = True
                        Else
                            var1 = 0
                            boolIgnoreFC = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLPSL", CompareMethod.Text) = 0 Then
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            boolPSL = True
                        Else
                            var1 = 0
                            boolPSL = False
                        End If

                    ElseIf StrComp(c.ColumnName, "CHARSTPAGE", CompareMethod.Text) = 0 Then
                        LcharSTPage = var1


                    ElseIf StrComp(c.ColumnName, "BOOLRECSIGFIG", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLRECSIGFIG)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLRECSIGFIG = True
                        Else
                            var1 = 0
                            BOOLRECSIGFIG = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLSD2", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLSD2)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLSD2 = True
                            gSDMax = 2
                        Else
                            var1 = 0
                            BOOLSD2 = False
                            gSDMax = 3
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLDIFFCOLSTATS", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLDIFFCOLSTATS)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLDIFFCOLSTATS = True
                        Else
                            var1 = 0
                            BOOLDIFFCOLSTATS = False
                        End If

                    ElseIf StrComp(c.ColumnName, "CHARCAPTIONFOLLOW", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, "Tab")
                        gstrCAPTIONFOLLOW = var1


                    ElseIf StrComp(c.ColumnName, "BOOLUSERSD", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLUSERSD)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLUSERSD = True
                        Else
                            var1 = 0
                            BOOLUSERSD = False
                        End If

                    ElseIf StrComp(c.ColumnName, "BOOLTABLELABELSECTION", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLTABLELABELSECTION)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLTABLELABELSECTION = True
                        Else
                            var1 = 0
                            BOOLTABLELABELSECTION = False
                        End If

                    ElseIf StrComp(c.ColumnName, "NUMTABLEFONTSIZE", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, NUMTABLEFONTSIZE)
                        NUMTABLEFONTSIZE = var1

                        '20190108 LEE:
                    ElseIf StrComp(c.ColumnName, "BOOLCALIBRTABLETITLE", CompareMethod.Text) = 0 Then
                        var1 = NZ(var1, BOOLCALIBRTABLETITLE)
                        If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                            var1 = -1
                            BOOLCALIBRTABLETITLE = True
                        Else
                            var1 = 0
                            BOOLCALIBRTABLETITLE = False
                        End If
                    End If

                    'drows(0).Item(c.ColumnName) = var1
                    'var1 = NZ(var1, 0)
                    Try
                        drows(0).Item(c.ColumnName) = var1
                        'check for LSigFigs LDateFormat
                    Catch ex As Exception
                        '20170913 LEE: This is problematic. Column data types in tblData are number, text and datetime
                        'instead, leave as null
                        'drows(0).Item(c.ColumnName) = ""
                    End Try


                End If 'end boolTable

                '****
                If boolRoundConv Then

                    If StrComp(c.ColumnName, "BOOLROUNDFIVEEVEN", CompareMethod.Text) = 0 Then

                        var2 = gboolRoundFiveEven

                    ElseIf StrComp(c.ColumnName, "BOOLROUNDFIVEAWAY", CompareMethod.Text) = 0 Then

                        var2 = gboolRoundFiveAway

                    ElseIf StrComp(c.ColumnName, "BOOLCRITFULLPREC", CompareMethod.Text) = 0 Then

                        var2 = gboolCritFullPrec

                    ElseIf StrComp(c.ColumnName, "BOOLCRITROUNDED", CompareMethod.Text) = 0 Then

                        var2 = gboolCritRounded

                    ElseIf StrComp(c.ColumnName, "BOOLMEANFULLPREC", CompareMethod.Text) = 0 Then

                        var2 = gboolMeanFullPrec

                    ElseIf StrComp(c.ColumnName, "BOOLMEANROUNDED", CompareMethod.Text) = 0 Then

                        var2 = gboolMeanRounded

                    End If

                    If var2 Then
                        var1 = -1
                    Else
                        var1 = 0
                    End If

                    Try
                        drows(0).Item(c.ColumnName) = var1
                        'check for LSigFigs LDateFormat
                    Catch ex As Exception
                        '20170913 LEE: This is problematic. Column data types in tblData are number, text and datetime
                        'instead, set as 0
                        drows(0).Item(c.ColumnName) = 0
                    End Try

                End If
                '*****

                If boolChk Then
                    var1 = 0
                    Select Case c.ColumnName
                        Case "BOOLQAEVENTBORDER"
                            If frmH.chkQAEventBorder.Checked Then
                                var1 = -1
                                BOOLQAEVENTBORDER = True
                            Else
                                var1 = 0
                                BOOLQAEVENTBORDER = False
                            End If
                    End Select
                    drows(0).Item(c.ColumnName) = var1
                End If

                If boolcbx Then
                    'var1 = NZ(drows(0).Item(c.ColumnName), "")
                    var1 = contr.Text
                    If Len(var1) = 0 Then
                        var2 = 0 '"[None]"
                    Else
                        str2 = "id_tblDropdownBoxName = " & strFld & " AND charValue = '" & var1 & "'"
                        dRows1 = tbl.Select(str2)
                        If dRows1.Length = 0 Then
                            var2 = 0
                        Else
                            var2 = dRows1(0).Item("id_tblDropdownBoxContent")
                        End If
                    End If
                    drows(0).Item(c.ColumnName) = var2
                End If
                If boolCorp Then
                    'tbl = tblCorporateAddresses
                    tbl = tblCorporateNickNames
                    strFld = "charNickname"
                    'strFld1 = "id_tblCorporateAddresses"
                    strFld1 = "id_tblCorporateNickNames"
                    'var1 = NZ(drows(0).Item(c.ColumnName), "")
                    var1 = contr.Text
                    If Len(var1) = 0 Then
                        'var2 = "[None]"
                        var2 = 0
                    Else
                        str2 = strFld & " = '" & var1 & "'"
                        dRows1 = tbl.Select(str2, "id_tblCorporateNickNames ASC")
                        If dRows1.Length = 0 Then
                            'var2 = "[None]"
                            var2 = 0
                        Else
                            var2 = dRows1(0).Item(strFld1)
                        End If
                    End If
                    'contr.Text = var2
                    drows(0).Item(c.ColumnName) = var2

                End If
                If boolEntireR Then
                    'make sure user has permission to modify report body section
                    Dim boolA As Short
                    boolA = BOOLREPORTBODYSECTIONS
                    If boolA = -1 Then
                        drows(0).Item(c.ColumnName) = numEntireR
                    End If

                End If

next1:

            Next
            drows(0).EndEdit()
        End If

        Dim dvCheck As System.Data.DataView = New DataView(tblData)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblData)

            If boolGuWuOracle Then
                Try
                    ta_tblData.Update(tblData)
                Catch ex As DBConcurrencyException
                    '''msgbox("aaData Tab: " & ex.Message)
                    'ds2005.TBLDATA.Merge('ds2005.TBLDATA, True)
                End Try

                'sometimes tbleData will not save. Try running it twice
                Try
                    ta_tblData.Update(tblData)
                Catch ex As DBConcurrencyException
                    '''msgbox("aaData Tab: " & ex.Message)
                    'ds2005.TBLDATA.Merge('ds2005.TBLDATA, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblDataAcc.Update(tblData)
                Catch ex As DBConcurrencyException
                    '''msgbox("aaData Tab: " & ex.Message)
                    'ds2005Acc.TBLDATA.Merge('ds2005Acc.TBLDATA, True)
                End Try

                'sometimes tbleData will not save. Try running it twice
                Try
                    ta_tblDataAcc.Update(tblData)
                Catch ex As DBConcurrencyException
                    '''msgbox("aaData Tab: " & ex.Message)
                    'ds2005Acc.TBLDATA.Merge('ds2005Acc.TBLDATA, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblDataSQLServer.Update(tblData)
                Catch ex As DBConcurrencyException
                    '''msgbox("aaData Tab: " & ex.Message)
                    'ds2005Acc.TBLDATA.Merge('ds2005Acc.TBLDATA, True)
                End Try

                'sometimes tbleData will not save. Try running it twice
                Try
                    ta_tblDataSQLServer.Update(tblData)
                Catch ex As DBConcurrencyException
                    '''msgbox("aaData Tab: " & ex.Message)
                    'ds2005Acc.TBLDATA.Merge('ds2005Acc.TBLDATA, True)
                End Try
            End If

        End If


        frmH.dgvDataCompany.AutoResizeColumns()
        frmH.dgvStudyConfig.AutoResizeColumns()

        'now save Analyte Sort
        Call SaveAnalyteSort()

    End Sub

    Sub SaveAnalyteSort()

        ''debug
        'Dim dtbl As DataTable = TBLSTUDYDOCANALYTES
        'Dim int1 As Int16
        'Dim int2 As Int16
        'Dim int3 As Int16

        'Dim rows1() As DataRow = dtbl.Select("", "", DataViewRowState.CurrentRows)
        'Dim rows2() As DataRow = dtbl.Select("", "", DataViewRowState.ModifiedCurrent)
        'Dim rows3() As DataRow = dtbl.Select("", "", DataViewRowState.ModifiedOriginal)

        'int1 = rows1.Length
        'int2 = rows2.Length
        'int3 = rows3.Length

        'Dim Count1 As Int16
        'Dim var1, var2, var3
        'Console.WriteLine("StartCurrent")
        'For Count1 = 0 To int2 - 1
        '    var1 = rows2(Count1).Item("AnalyteDescription")
        '    var2 = rows2(Count1).Item("INTORDER")
        '    Console.WriteLine(var1.ToString)
        '    Console.WriteLine(var2.ToString)
        '    var2 = var2
        'Next
        'Console.WriteLine("EndCurrent")
        'Console.WriteLine("StartOriginal")
        'For Count1 = 0 To int2 - 1
        '    var1 = rows3(Count1).Item("AnalyteDescription")
        '    var2 = rows3(Count1).Item("INTORDER")
        '    Console.WriteLine(var1.ToString)
        '    Console.WriteLine(var2.ToString)
        '    var2 = var2
        'Next
        'Console.WriteLine("EndOriginal")

        Call SaveAnalyteGroups()

        Call FillAuditTrailTemp(TBLSTUDYDOCANALYTES)

        If boolGuWuOracle Then
            Try
                'ta_TBLSTUDYDOCANALYTES.Update(TBLSTUDYDOCANALYTES)
            Catch ex As DBConcurrencyException
                'ds2005.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005.TBLCONTRIBUTINGPERSONNEL, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLSTUDYDOCANALYTESAcc.Update(TBLSTUDYDOCANALYTES)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005Acc.TBLCONTRIBUTINGPERSONNEL, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLSTUDYDOCANALYTESSQLSERVER.Update(TBLSTUDYDOCANALYTES)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005Acc.TBLCONTRIBUTINGPERSONNEL, True)
            End Try
        End If

    End Sub

    Sub FillDataTabData(ByVal boolFromReset As Boolean)

        Dim str1 As String
        Dim str2 As String
        Dim drows() As DataRow
        Dim var1, var2, var3
        Dim c As DataColumn
        Dim tbl As System.Data.DataTable
        Dim boolcbx As Boolean
        Dim boolCorp As Boolean
        Dim boolTable As Boolean
        Dim boolChk As Boolean
        Dim strFld As String
        Dim strFld1 As String
        Dim contr As Control
        Dim drows2() As DataRow
        Dim int1 As Short
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim Count1 As Short
        Dim ct1 As Short
        Dim dvD As System.Data.DataView
        Dim dvD1 As System.Data.DataView
        Dim boolEntireR As Boolean
        Dim intRow As Short
        Dim strF As String
        Dim strS As String
        Dim boolIsBool As Boolean
        Dim boolDD As Boolean
        Dim boolRoundConv As Boolean = False

        boolFromCD = True

        If id_tblStudies = 0 Then
            Exit Sub
        End If

        boolStopCBX = True

        str1 = "id_tblStudies = " & id_tblStudies
        drows = tblData.Select(str1, "id_tblData ASC")
        ct1 = drows.Length
        dvD = frmH.dgvDataCompany.DataSource
        dvD1 = frmH.dgvStudyConfig.DataSource

        If ct1 = 0 Then 'no records to retrieve

            Try
                'populate default values in dgvData
                'int1 = FindRowDV("Data Significant Figures", dvD1)
                int1 = FindRowDV("Data Sig Figs/Decimals", dvD1)
                dvD1(int1).BeginEdit()
                dvD1(int1).Item("Value") = GSigFig
                dvD1(int1).EndEdit()

                int1 = FindRowDV("Table Date Format", dvD1)
                dvD1(int1).BeginEdit()
                dvD1(int1).Item("Value") = GDateFormat
                dvD1(int1).EndEdit()

                int1 = FindRowDV("Text Date Format", dvD1)
                dvD1(int1).BeginEdit()
                dvD1(int1).Item("Value") = GTextDateFormat
                dvD1(int1).EndEdit()

                int1 = FindRowDV("Data Decimal Places", dvD1)
                dvD1(int1).BeginEdit()
                dvD1(int1).Item("Value") = GDec
                dvD1(int1).EndEdit()

                int1 = FindRowDV("Time Zone", dvD1)
                dvD1(int1).BeginEdit()
                dvD1(int1).Item("Value") = GTimeZone
                dvD1(int1).EndEdit()

                int1 = FindRowDV("Data: Use Sig Figs, not Decimals", dvD1)
                dvD1(int1).BeginEdit()
                If boolGUseSigFigs Then
                    dvD1(int1).Item("Value") = "TRUE"
                Else
                    dvD1(int1).Item("Value") = "FALSE"
                End If
                dvD1(int1).EndEdit()
                boolLUseSigFigs = boolGUseSigFigs

                intRow = int1
                intRow = FindRowDV("Data: Use Conc Special Rounding", dvD1)
                dvD1(intRow).BeginEdit()
                If GboolWyethRounding Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()
                LboolWyethRounding = GboolWyethRounding 'establish local sigfig

                'intRow = FindRowDV("Regr Const Sig Figs", dvD1)
                intRow = FindRowDV("Regr Const Sig Figs/Decimals", dvD1)
                dvD1(intRow).BeginEdit()
                dvD1(intRow).Item("Value") = GRegrSigFigs
                dvD1(intRow).EndEdit()
                LRegrSigFigs = GRegrSigFigs

                'intRow = FindRowDV("Regr R2 Sig Figs", dvD1)
                intRow = FindRowDV("Regr R2 Sig Figs/Decimals", dvD1)
                dvD1(intRow).BeginEdit()
                dvD1(intRow).Item("Value") = GR2SigFigs
                dvD1(intRow).EndEdit()
                LR2SigFigs = GR2SigFigs

                '***

                'intRow = FindRowDV("Peak Area Significant Figures", dvD1)
                intRow = FindRowDV("Peak Area Sig Figs/Decimals", dvD1)
                dvD1(intRow).BeginEdit()
                dvD1(intRow).Item("Value") = GSigFigArea
                dvD1(intRow).EndEdit()
                LSigFigArea = GSigFigArea 'establish local 

                'intRow = FindRowDV("Peak Area Decimal Places", dvD1)
                'dvD1(intRow).BeginEdit()
                'dvD1(intRow).Item("Value") = GDecArea
                'dvD1(intRow).EndEdit()
                GDecArea = GSigFigArea
                LDecArea = GDecArea 'establish local 

                intRow = FindRowDV("Peak Areas: Use Sig Figs, not Decimals", dvD1)
                dvD1(intRow).BeginEdit()
                If boolGUseSigFigsArea Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()
                boolLUseSigFigsArea = boolGUseSigFigsArea 'establish local sigfig

                intRow = FindRowDV("Peak Areas: Use Conc Special Rounding", dvD1)
                dvD1(intRow).BeginEdit()
                If GboolWyethRoundingArea Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()
                LboolWyethRoundingArea = GboolWyethRoundingArea 'establish local 


                intRow = FindRowDV("Peak Area Ratio: Use Sig Figs, not Decimals", dvD1)
                dvD1(intRow).BeginEdit()
                If boolGUseSigFigsArea Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()
                boolLUseSigFigsAreaRatio = boolGUseSigFigsAreaRatio 'establish local sigfig

                intRow = FindRowDV("Peak Area Ratio: Use Conc Special Rounding", dvD1)
                dvD1(intRow).BeginEdit()
                If GboolWyethRoundingArea Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()
                LboolWyethRoundingAreaRatio = GboolWyethRoundingAreaRatio 'establish local 



                int1 = FindRowDV("Regression and R2: Use Sig Figs, not Decimals", dvD1)
                dvD1(int1).BeginEdit()
                If boolGUseSigFigsRegr Then
                    dvD1(int1).Item("Value") = "TRUE"
                Else
                    dvD1(int1).Item("Value") = "FALSE"
                End If
                dvD1(int1).EndEdit()
                boolLUseSigFigsRegr = boolGUseSigFigsRegr

                intRow = FindRowDV("Regression and R2: Use Sci. Notation", dvD1)
                dvD1(intRow).BeginEdit()
                dvD1(intRow).Item("Value") = boolGUseRegrSciNot
                dvD1(intRow).EndEdit()
                boolLUseRegrSciNot = boolLUseRegrSciNot  'establish local 


                'intRow = FindRowDV("Regression and R2 Decimal Places", dvD1)
                'dvD1(intRow).BeginEdit()
                'dvD1(intRow).Item("Value") = GRegrDec
                'dvD1(intRow).EndEdit()
                LRegrDec = GRegrDec 'establish local 
                LRegrDec = LRegrSigFigs

                '***

                intRow = FindRowDV("Alternate Calibr/QC Std Units", dvD1)
                dvD1(intRow).BeginEdit()
                dvD1(intRow).Item("Value") = System.DBNull.Value
                dvD1(intRow).EndEdit()

                'Default # of Decimals for QC Stats
                intRow = FindRowDV("QC Stats % Decimal Places", dvD1)
                dvD1(intRow).BeginEdit()
                dvD1(intRow).Item("Value") = gintQCDec
                dvD1(intRow).EndEdit()

                intRow = FindRowDV("Enable StudyDoc Exclude Samples feature", dvD1)
                dvD1(intRow).BeginEdit()
                If gAllowExclSamples Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()

                intRow = FindRowDV("Enable StudyDoc Acceptance Crit. feature", dvD1)
                dvD1(intRow).BeginEdit()
                If gAllowExclSamples Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()

                intRow = FindRowDV("Make hyperlinks and TOC blue font color", dvD1)
                dvD1(intRow).BeginEdit()
                If boolBLUEHYPERLINK Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()

                'intRow = FindRowDV("Format table anomalies with red bold font", dvD1)
                'dvD1(intRow).BeginEdit()
                intRow = FindRowDV("Format table anomalies with red bold font", dvD1)
                dvD1(intRow).BeginEdit()
                If boolRedBoldFont Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()

                'intRow = FindRowDV("Use SigFigs for Recovery/MatrixFactor values", dvD1)
                '20181109 LEE:
                intRow = FindRowDV("Use SigFigs for Recovery values", dvD1)
                dvD1(intRow).BeginEdit()
                If BOOLRECSIGFIG Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If
                dvD1(intRow).EndEdit()

                intRow = FindRowDV("Allow StdDev calculation if n = 2", dvD1)
                dvD1(intRow).BeginEdit()
                If BOOLSD2 Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If

                intRow = FindRowDV("If %Accuracy Column is displayed, then report average in Statistics section", dvD1)
                dvD1(intRow).BeginEdit()
                If BOOLDIFFCOLSTATS Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If

                intRow = FindRowDV("Character following table/figure/appendix caption", dvD1)
                dvD1(intRow).BeginEdit()
                '20171117 LEE: Possible NULL problem
                'dvD1(intRow).Item("Value") = gstrCAPTIONFOLLOW
                dvD1(intRow).Item("Value") = NZ(gstrCAPTIONFOLLOW, "Tab")


                intRow = FindRowDV("Use %RSD label instead of %CV", dvD1)
                dvD1(intRow).BeginEdit()
                If BOOLUSERSD Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If

                intRow = FindRowDV("Add chapter number to table caption label", dvD1)
                dvD1(intRow).BeginEdit()
                If BOOLTABLELABELSECTION Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If

                intRow = FindRowDV("Table font size (0 to use Normal style font size)", dvD1)
                dvD1(intRow).BeginEdit()
                dvD1(intRow).Item("Value") = 0

                '20190108 LEE:
                intRow = FindRowDV("Include Calibr Range in table title if multi-calibr range", dvD1)
                dvD1(intRow).BeginEdit()
                If BOOLCALIBRTABLETITLE Then
                    dvD1(intRow).Item("Value") = "TRUE"
                Else
                    dvD1(intRow).Item("Value") = "FALSE"
                End If

                dvD1(intRow).EndEdit()


            Catch ex As Exception


            End Try


            BOOLQAEVENTBORDER = False

            GoTo end1

        End If

        'dv = tblDataTableRowTitles.DefaultView

        strF = "BOOLINCLUDE <> 0 "
        strS = "ID_tblDataTableRowTitles ASC"
        dv = New DataView(tblDataTableRowTitles, strF, strS, DataViewRowState.CurrentRows)

        int1 = tblData.Columns.Count

        Dim dRows1() As DataRow
        Count1 = 0
        '''''''''''console.writeline("Start tblData columns...")

        Dim boolSkip As Boolean = False

        Dim ctCol As Int16 = tblData.Columns.Count 'debug

        For Each c In tblData.Columns

            boolcbx = False
            boolCorp = False
            boolTable = False
            boolEntireR = False
            boolChk = False
            boolIsBool = False 'this is for dgvStudyConfig
            boolDD = False 'this is for dgvStudyConfig
            boolRoundConv = False

            boolSkip = False

            Count1 = Count1 + 1

            If Count1 = ctCol Then
                var1 = var1'debug
            End If

            '''''''''''console.writeline(c.ColumnName)

            Select Case c.ColumnName
                Case "ID_TBLDATA"
                Case "ID_TBLSTUDIES"
                Case "ID_TBLASSAYTECHNIQUE" 'ID_TBLASSAYTECHNIQUE
                    tbl = tblDropdownBoxContent
                    strFld = "3"
                    strFld1 = c.ColumnName
                    contr = frmH.cbxAssayTechnique
                    boolcbx = True
                Case "ID_TBLANTICOAGULANT" 'ID_TBLANTICOAGULANT
                    tbl = tblDropdownBoxContent
                    strFld = "1"
                    strFld1 = c.ColumnName
                    contr = frmH.cbxAnticoagulant
                    boolcbx = True
                Case "CHARCORPORATESTUDYID" 'CHARCORPORATESTUDYID
                    boolTable = True
                Case "CHARPROTOCOLNUMBER" 'CHARPROTOCOLNUMBER
                    boolTable = True
                Case "ID_TBLVOLUMEUNITS" 'ID_TBLVOLUMEUNITS
                    'deprecated. now in tblMethValidation
                    tbl = tblDropdownBoxContent
                    strFld = "10"
                    strFld1 = "id_tblVolumeUnits"
                    contr = frmH.cbxSampleSizeUnits
                    boolcbx = True
                Case "ID_TBLTEMPERATURES" 'ID_TBLTEMPERATURES
                    'deprecated. now in tblMethValidation
                    tbl = tblDropdownBoxContent
                    strFld = "9"
                    strFld1 = "id_tblTemperatures"
                    contr = frmH.cbxSampleStorageTemp
                    boolcbx = True
                Case "ID_SUBMITTEDBY" 'ID_SUBMITTEDBY
                    contr = frmH.cbxSubmittedBy
                    boolCorp = True
                Case "ID_SUBMITTEDTO"
                    contr = frmH.cbxSubmittedTo
                    boolCorp = True
                Case "ID_INSUPPORTOF"
                    contr = frmH.cbxInSupportOf
                    boolCorp = True
                Case "CHARDATAARCHIVALLOCATION" 'CHARDATAARCHIVALLOCATION
                    boolTable = True

                Case "CHARSPONSORSTUDYNUMBER" 'CHARSPONSORSTUDYNUMBER
                    boolTable = True
                Case "CHARSPONSORSTUDYTITLE" 'CHARSPONSORSTUDYTITLE
                    boolTable = True
                Case "NUMSIGFIGS" '
                    boolTable = True
                Case "CHARDATEFORMAT"
                    boolTable = True
                Case "CHARTEXTDATEFORMAT"
                    boolTable = True
                Case "NUMDECIMALS" '
                    boolTable = True
                Case "BOOLUSESIGFIGS" '
                    boolTable = True
                Case "CHARTIMEZONE" 'CHARTIMEZONE
                    boolTable = True
                Case "CHAROUTLIERMETHOD" 'CHAROUTLIERMETHOD
                    boolTable = True
                Case "BOOLENTIREREPORT" '
                    boolEntireR = True
                Case "BOOLUSESPECRND"
                    boolTable = True
                Case "NUMREGRSIGFIGS"
                    boolTable = True
                Case "NUMR2SIGFIGS"
                    boolTable = True
                Case "CHARUNITS"
                    boolTable = True
                Case "INTQCPERCDECPLACES"
                    boolTable = True
                Case "BOOLQAEVENTBORDER"
                    boolChk = True
                Case "BOOLALLOWEXCLSAMPLES"
                    boolTable = True
                Case "BOOLALLOWGUWUACCCRIT"
                    boolTable = True
                Case "DTSTUDYSTARTDATE"
                    boolTable = True
                Case "DTSTUDYENDDATE"
                    boolTable = True
                Case "INTCOMMAFORMAT"
                    boolTable = True
                Case "BOOLBLUEHYPERLINK"
                    boolTable = True
                Case "BOOLREDBOLDFONT"
                    boolTable = True



                Case "NUMSIGFIGSAREA" '
                    boolTable = True
                Case "NUMDECIMALSAREA" '
                    boolTable = True
                Case "BOOLUSESIGFIGSAREA" '
                    boolTable = True
                Case "BOOLUSESPECRNDAREA"
                    boolTable = True


                Case "NUMSIGFIGSAREARATIO" '
                    boolTable = True
                Case "NUMDECIMALSAREARATIO" '
                    boolTable = True
                Case "BOOLUSESIGFIGSAREARATIO" '
                    boolTable = True
                Case "BOOLUSESPECRNDAREARATIO"
                    boolTable = True


                Case "BOOLUSESIGFIGSREGR" '
                    boolTable = True
                Case "NUMREGRDEC" '
                    boolTable = True

                Case "BOOLUSEREGRSCINOT"
                    boolTable = True

                Case "BOOLNOMCONCPAREN"
                    boolTable = True
                Case "CHARSTPAGE"
                    boolTable = True

                Case "BOOLTABLEDTTIMESTAMP"
                    boolTable = True

                Case "BOOLFOOTNOTEQCMEAN"
                    boolTable = True
                Case "BOOLFLIPHEADER"
                    boolTable = True

                Case "BOOLQCNA"
                    boolTable = True

                Case "BOOLBQL"
                    boolTable = True

                    boolSkip = True 'deprecate

                Case "CHARBQL"
                    boolTable = True

                Case "BOOLIGNOREFC"
                    boolTable = True

                Case "BOOLPSL"
                    boolTable = True



                Case "BOOLROUNDFIVEEVEN"
                    boolRoundConv = True

                Case "BOOLROUNDFIVEAWAY"
                    boolRoundConv = True

                Case "BOOLCRITFULLPREC"
                    boolRoundConv = True

                Case "BOOLCRITROUNDED"
                    boolRoundConv = True

                Case "BOOLMEANFULLPREC"
                    boolRoundConv = True

                Case "BOOLMEANROUNDED"
                    boolRoundConv = True


                Case "CHARCAPTIONTRAILER" '
                    boolTable = True

                Case "BOOLRECSIGFIG" '
                    boolTable = True

                Case "BOOLSD2" '
                    boolTable = True

                Case "BOOLDIFFCOLSTATS" '
                    boolTable = True

                Case "CHARCAPTIONFOLLOW" '
                    boolTable = True

                Case "BOOLUSERSD" '
                    boolTable = True

                Case "BOOLTABLELABELSECTION" '
                    boolTable = True

                Case "NUMTABLEFONTSIZE" '
                    boolTable = True

                    '20190108 LEE:
                Case "BOOLCALIBRTABLETITLE" '
                    boolTable = True

            End Select

            If InStr(1, c.ColumnName, "BOOL", CompareMethod.Text) > 0 Then
                boolIsBool = True
            End If

            Dim strTN As String
            'Dim strF As String
            Dim rowsAA() As DataRow

            If boolSkip Then
                GoTo skip1
            End If

            If boolTable Then

                If InStr(1, c.ColumnName, "BOOL", CompareMethod.Text) > 0 Or InStr(1, c.ColumnName, "NUM", CompareMethod.Text) > 0 Or InStr(1, c.ColumnName, "INT", CompareMethod.Text) > 0 Then
                    var1 = NZ(drows(0).Item(c.ColumnName), 0)
                Else
                    var1 = NZ(drows(0).Item(c.ColumnName), "[None]")
                End If


                str1 = "charDataTableName = 'tblCompanyData' and charTableRefColumnName = '" & c.ColumnName & "'"
                'get Tab type
                dv.RowFilter = str1
                var2 = NZ(dv.Item(0).Item("charRowName"), "")
                'int1 = FindRow(var2, tblCompanyData, "Item")

                strF = "Item = '" & var2 & "'"
                strTN = GetTabColumn(c.ColumnName)

                Erase rowsAA
                rowsAA = tblCompanyData.Select(strF)
                If rowsAA.Length > 0 Then

                    rowsAA(0).BeginEdit()
                    rowsAA(0).Item("charTab") = strTN
                    Try
                        rowsAA(0).Item("boolisBool") = boolIsBool
                    Catch ex As Exception
                        Dim varAAA
                        varAAA = ex.Message
                        varAAA = var1
                    End Try

                    Select Case strTN
                        Case "Data"

                            If StrComp(var2, "Study Start Date", CompareMethod.Text) = 0 Then
                                If IsDBNull(var1) Or Len(NZ(var1, "")) = 0 Or StrComp(CStr(NZ(var1, "")), "[None]", CompareMethod.Text) = 0 Then
                                    rowsAA(0).Item(1) = DBNull.Value
                                Else
                                    If IsDate(var1) Then
                                        Try
                                            rowsAA(0).Item(1) = Format(var1, LDateFormat)
                                        Catch ex As Exception
                                            rowsAA(0).Item(1) = Format(var1, GDateFormat)
                                        End Try
                                    Else
                                        rowsAA(0).Item(1) = DBNull.Value
                                    End If
                                End If
                            ElseIf StrComp(var2, "Study End Date", CompareMethod.Text) = 0 Then
                                If IsDBNull(var1) Or Len(NZ(var1, "")) = 0 Or StrComp(CStr(NZ(var1, "")), "[None]", CompareMethod.Text) = 0 Then
                                    rowsAA(0).Item(1) = DBNull.Value
                                Else
                                    If IsDate(var1) Then
                                        Try
                                            rowsAA(0).Item(1) = Format(var1, LDateFormat)
                                        Catch ex As Exception
                                            rowsAA(0).Item(1) = Format(var1, GDateFormat)
                                        End Try
                                    Else
                                        rowsAA(0).Item(1) = DBNull.Value
                                    End If
                                End If
                            Else
                                rowsAA(0).Item(1) = var1
                            End If

                        Case "Config"

                            If StrComp(var2, "Data: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If
                            ElseIf StrComp(var2, "Data: Use Conc Special Rounding", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                                '******
                            ElseIf StrComp(var2, "Peak Areas: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If
                            ElseIf StrComp(var2, "Peak Areas: Use Conc Special Rounding", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Peak Area Ratio: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If
                            ElseIf StrComp(var2, "Peak Area Ratio: Use Conc Special Rounding", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If


                            ElseIf StrComp(var2, "Regression and R2: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Regression and R2: Use Sci. Notation", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                                '******

                            ElseIf StrComp(var2, "Enable StudyDoc Exclude Samples feature", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If
                            ElseIf StrComp(var2, "Enable StudyDoc Acceptance Crit. feature", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If
                            ElseIf StrComp(var2, "Make hyperlinks and TOC blue font color", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If
                            ElseIf StrComp(var2, "Format table anomalies with red bold font", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Place Nominal Concentrations in parentheses", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Add a date/time stamp on tables", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Footnote QC Means that exceed acceptance criteria", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Header/footer in right/left margin on landscape page", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Enter NA for non-entry QC or Calibr Std levels", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Allow StdDev calculation if n = 2", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "If %Accuracy Column is displayed, then report average in Statistics section", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Use BQL/AQL vs BLQ/ALQ vs LLOQ/ULOQ", CompareMethod.Text) = 0 Then
                                rowsAA(0).Item(1) = NZ(var1, "BQL/AQL")
                          

                            ElseIf StrComp(var2, "Ignore Table-Specific Field Code generation", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Enable Page-Specific Legends", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                                '20181109 LEE:
                                'ElseIf StrComp(var2, "Use SigFigs for Recovery/MatrixFactor values", CompareMethod.Text) = 0 Then
                            ElseIf StrComp(var2, "Use SigFigs for Recovery values", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Appendix/Figure/Table Caption Trailer", CompareMethod.Text) = 0 Then
                                If StrComp(var1, "[NONE]", CompareMethod.Text) = 0 Then
                                    rowsAA(0).Item(1) = ""
                                Else
                                    rowsAA(0).Item(1) = var1
                                End If

                            ElseIf StrComp(var2, "Character following table/figure/appendix caption", CompareMethod.Text) = 0 Then
                                rowsAA(0).Item(1) = NZ(var1, "Tab")
                                gstrCAPTIONFOLLOW = var1

                            ElseIf StrComp(var2, "Use %RSD label instead of %CV", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Add chapter number to table caption label", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            ElseIf StrComp(var2, "Table font size (0 to use Normal style font size)", CompareMethod.Text) = 0 Then
                                rowsAA(0).Item(1) = NZ(var1, NZ(NUMTABLEFONTSIZE, 0))

                                '20190108 LEE
                            ElseIf StrComp(var2, "Include Calibr Range in table title if multi-calibr range", CompareMethod.Text) = 0 Then
                                If var1 = 0 Then
                                    rowsAA(0).Item(1) = "FALSE"
                                Else
                                    rowsAA(0).Item(1) = "TRUE"
                                End If

                            Else
                                rowsAA(0).Item(1) = var1
                            End If
                    End Select

                    rowsAA(0).EndEdit()

                End If


                'dvD(int1).BeginEdit()

                'dvD(int1).EndEdit()
                If StrComp(c.ColumnName, "NUMSIGFIGS", CompareMethod.Text) = 0 Then
                    'record lsigfig
                    LSigFig = CInt(var1)
                ElseIf StrComp(c.ColumnName, "CHARDATEFORMAT", CompareMethod.Text) = 0 Then
                    LDateFormat = NZ(var1, GDateFormat)
                ElseIf StrComp(c.ColumnName, "CHARTEXTDATEFORMAT", CompareMethod.Text) = 0 Then
                    LTextDateFormat = NZ(var1, GTextDateFormat)
                    'seems that a YYYY gets in there somehow
                    LTextDateFormat = Replace(LTextDateFormat, "YYYY", "yyyy", 1, -1, CompareMethod.Binary)

                ElseIf StrComp(c.ColumnName, "NUMDECIMALS", CompareMethod.Text) = 0 Then
                    'var1 = NZ(var1, GDec)
                    'LDec = var1
                    var1 = NZ(var1, GSigFig)
                    LDec = LSigFig ' var1
                ElseIf StrComp(c.ColumnName, "CHARTIMEZONE", CompareMethod.Text) = 0 Then
                    var1 = NZ(var1, GTimeZone)
                    LTimeZone = var1
                ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGS", CompareMethod.Text) = 0 Then
                    'var1 = NZ(var1, boolGUseSigFigs)
                    If var1 = 0 Then
                        boolLUseSigFigs = False
                    Else
                        boolLUseSigFigs = True
                    End If
                ElseIf StrComp(c.ColumnName, "BOOLUSESPECRND", CompareMethod.Text) = 0 Then
                    'var1 = NZ(var1, boolGUseSigFigs)
                    If var1 = 0 Then
                        LboolWyethRounding = False
                    Else
                        LboolWyethRounding = True
                    End If
                ElseIf StrComp(c.ColumnName, "NUMREGRSIGFIGS", CompareMethod.Text) = 0 Then
                    var1 = NZ(var1, GRegrSigFigs)
                    LRegrSigFigs = var1
                    LRegrDec = LRegrSigFigs
                ElseIf StrComp(c.ColumnName, "NUMR2SIGFIGS", CompareMethod.Text) = 0 Then
                    var1 = NZ(var1, GR2SigFigs)
                    LR2SigFigs = var1
                    'lr2dec = LR2SigFigs

                    '******

                ElseIf StrComp(c.ColumnName, "NUMSIGFIGSAREA", CompareMethod.Text) = 0 Then
                    var1 = NZ(var1, GSigFigArea)
                    LSigFigArea = var1
                    LDecArea = LSigFigArea
                ElseIf StrComp(c.ColumnName, "NUMDECIMALSAREA", CompareMethod.Text) = 0 Then
                    'var1 = NZ(var1, GDecArea)
                    'LDecArea = var1
                ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSAREA", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolLUseSigFigsArea = False
                    Else
                        boolLUseSigFigsArea = True
                    End If
                ElseIf StrComp(c.ColumnName, "BOOLUSESPECRNDAREA", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        LboolWyethRoundingArea = False
                    Else
                        LboolWyethRoundingArea = True
                    End If

                    '*****

                ElseIf StrComp(c.ColumnName, "NUMSIGFIGSAREARATIO", CompareMethod.Text) = 0 Then
                    var1 = NZ(var1, GSigFigAreaRatio)
                    LSigFigAreaRatio = var1
                    LDecAreaRatio = LSigFigAreaRatio
                    strAreaDecAreaRatio = GetAreaRatioDecStr()

                ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSAREARATIO", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolLUseSigFigsAreaRatio = False
                    Else
                        boolLUseSigFigsAreaRatio = True
                    End If
                ElseIf StrComp(c.ColumnName, "BOOLUSESPECRNDAREARATIO", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        LboolWyethRoundingAreaRatio = False
                    Else
                        LboolWyethRoundingAreaRatio = True
                    End If

                    '******

                    '*****'
                ElseIf StrComp(c.ColumnName, "BOOLRECSIGFIG", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLRECSIGFIG = False
                    Else
                        BOOLRECSIGFIG = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLSD2", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLSD2 = False
                        gSDMax = 3
                    Else
                        BOOLSD2 = True
                        gSDMax = 2
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLDIFFCOLSTATS", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLDIFFCOLSTATS = False
                    Else
                        BOOLDIFFCOLSTATS = True
                    End If


                ElseIf StrComp(c.ColumnName, "CHARCAPTIONFOLLOW", CompareMethod.Text) = 0 Then
                    var1 = NZ(var1, "Tab")
                    gstrCAPTIONFOLLOW = var1


                ElseIf StrComp(c.ColumnName, "BOOLUSERSD", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLUSERSD = False
                    Else
                        BOOLUSERSD = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLTABLELABELSECTION", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLTABLELABELSECTION = False
                    Else
                        BOOLTABLELABELSECTION = True
                    End If

                    '20190108 LEE:
                ElseIf StrComp(c.ColumnName, "BOOLCALIBRTABLETITLE", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLCALIBRTABLETITLE = False
                    Else
                        BOOLCALIBRTABLETITLE = True
                    End If

                ElseIf StrComp(c.ColumnName, "NUMTABLEFONTSIZE", CompareMethod.Text) = 0 Then
                    NUMTABLEFONTSIZE = var1

                ElseIf StrComp(c.ColumnName, "CHARCAPTIONTRAILER", CompareMethod.Text) = 0 Then

                    'user may enter blank
                    'dgv shows this as null
                    'must convert to zero-length string

                    'previous code may change "" to [NONE]

                    If StrComp(var1, "[NONE]", CompareMethod.Text) = 0 Then
                        var1 = ""
                    Else
                        var1 = NZ(var1, "")
                    End If


                    lcharCaptionTrailer = var1

                    '****


                ElseIf StrComp(c.ColumnName, "BOOLUSESIGFIGSREGR", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolLUseSigFigsRegr = False
                    Else
                        boolLUseSigFigsRegr = True
                    End If
                ElseIf StrComp(c.ColumnName, "BOOLUSEREGRSCINOT", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolLUseRegrSciNot = False
                    Else
                        boolLUseRegrSciNot = True
                    End If
                ElseIf StrComp(c.ColumnName, "NUMREGRDEC", CompareMethod.Text) = 0 Then
                    'var1 = NZ(var1, GRegrDec)
                    'LRegrDec = var1

                    '******

                ElseIf StrComp(c.ColumnName, "BOOLALLOWEXCLSAMPLES", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        LAllowExclSamples = False
                    Else
                        LAllowExclSamples = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLALLOWGUWUACCCRIT", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        LAllowGuWuAccCrit = False
                    Else
                        LAllowGuWuAccCrit = True
                    End If

                ElseIf StrComp(c.ColumnName, "INTCOMMAFORMAT", CompareMethod.Text) = 0 Then
                    gINTCOMMAFORMAT = NZ(var1, 0)

                ElseIf StrComp(c.ColumnName, "BOOLBLUEHYPERLINK", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolBLUEHYPERLINK = False
                    Else
                        boolBLUEHYPERLINK = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLREDBOLDFONT", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolRedBoldFont = False
                    Else
                        boolRedBoldFont = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLNOMCONCPAREN", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        LboolNomConcParen = False
                    Else
                        LboolNomConcParen = True
                    End If
                ElseIf StrComp(c.ColumnName, "CHARSTPAGE", CompareMethod.Text) = 0 Then
                    LcharSTPage = NZ(var1, LcharSTPage)

                ElseIf StrComp(c.ColumnName, "BOOLTABLEDTTIMESTAMP", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        LBOOLTABLEDTTIMESTAMP = False
                    Else
                        LBOOLTABLEDTTIMESTAMP = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLFOOTNOTEQCMEAN", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolFootNoteQCMean = False
                    Else
                        boolFootNoteQCMean = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLFLIPHEADER", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolFlipHeader = False
                    Else
                        boolFlipHeader = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLQCNA", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolQCNA = False
                    Else
                        boolQCNA = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLBQL", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolBQL = False
                    Else
                        boolBQL = True
                    End If

                ElseIf StrComp(c.ColumnName, "CHARBQL", CompareMethod.Text) = 0 Then

                    gstrBQL = NZ(var1, "BQL/AQL")

                ElseIf StrComp(c.ColumnName, "CHARCAPTIONFOLLOW", CompareMethod.Text) = 0 Then

                    gstrCAPTIONFOLLOW = NZ(var1, "Tab")


                ElseIf StrComp(c.ColumnName, "BOOLUSERSD", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLUSERSD = False
                    Else
                        BOOLUSERSD = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLTABLELABELSECTION", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLTABLELABELSECTION = False
                    Else
                        BOOLTABLELABELSECTION = True
                    End If

                    '20190108 LEE:
                ElseIf StrComp(c.ColumnName, "BOOLCALIBRTABLETITLE", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        BOOLCALIBRTABLETITLE = False
                    Else
                        BOOLCALIBRTABLETITLE = True
                    End If

                ElseIf StrComp(c.ColumnName, "NUMTABLEFONTSIZE", CompareMethod.Text) = 0 Then
                    NUMTABLEFONTSIZE = var1

                ElseIf StrComp(c.ColumnName, "BOOLIGNOREFC", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolIgnoreFC = False
                    Else
                        boolIgnoreFC = True
                    End If

                ElseIf StrComp(c.ColumnName, "BOOLPSL", CompareMethod.Text) = 0 Then
                    If var1 = 0 Then
                        boolPSL = False
                    Else
                        boolPSL = True
                    End If

                End If

            End If 'end boolTable

            '****
            If boolRoundConv Then

                var1 = NZ(drows(0).Item(c.ColumnName), 0)

                If StrComp(c.ColumnName, "BOOLROUNDFIVEEVEN", CompareMethod.Text) = 0 Then

                    If var1 = 0 Then
                        gboolRoundFiveEven = False
                    Else
                        gboolRoundFiveEven = True
                    End If

                    frmH.rbRoundFiveEven.Checked = gboolRoundFiveEven

                ElseIf StrComp(c.ColumnName, "BOOLROUNDFIVEAWAY", CompareMethod.Text) = 0 Then

                    If var1 = 0 Then
                        gboolRoundFiveAway = False
                    Else
                        gboolRoundFiveAway = True
                    End If
                    frmH.rbRoundFiveAway.Checked = gboolRoundFiveAway

                ElseIf StrComp(c.ColumnName, "BOOLCRITFULLPREC", CompareMethod.Text) = 0 Then

                    If var1 = 0 Then
                        gboolCritFullPrec = False
                    Else
                        gboolCritFullPrec = True
                    End If

                    frmH.rbCritFullPrec.Checked = gboolCritFullPrec

                ElseIf StrComp(c.ColumnName, "BOOLCRITROUNDED", CompareMethod.Text) = 0 Then

                    If var1 = 0 Then
                        gboolCritRounded = False
                    Else
                        gboolCritRounded = True
                    End If

                    frmH.rbCritRounded.Checked = gboolCritRounded

                ElseIf StrComp(c.ColumnName, "BOOLMEANFULLPREC", CompareMethod.Text) = 0 Then

                    If var1 = 0 Then
                        gboolMeanFullPrec = False
                    Else
                        gboolMeanFullPrec = True
                    End If

                    frmH.rbMeanFullPrec.Checked = gboolMeanFullPrec

                ElseIf StrComp(c.ColumnName, "BOOLMEANROUNDED", CompareMethod.Text) = 0 Then

                    If var1 = 0 Then
                        gboolMeanRounded = False
                    Else
                        gboolMeanRounded = True
                    End If

                    frmH.rbMeanRounded.Checked = gboolMeanRounded

                End If

            End If

            '*****

            If boolChk Then

                var1 = NZ(drows(0).Item(c.ColumnName), 0)
                Select Case c.ColumnName
                    Case "BOOLQAEVENTBORDER"
                        If var1 = -1 Then
                            frmH.chkQAEventBorder.Checked = True
                            BOOLQAEVENTBORDER = True
                        Else
                            frmH.chkQAEventBorder.Checked = False
                            BOOLQAEVENTBORDER = False
                        End If
                End Select
            End If
            If boolcbx Then
                var1 = NZ(drows(0).Item(c.ColumnName), "") 'drow hits on tblData
                If Len(var1) = 0 Then
                    var2 = "[None]"
                Else
                    str2 = "id_tblDropdownBoxContent = " & var1
                    dRows1 = tbl.Select(str2)
                    If dRows1.Length = 0 Then
                    Else
                        'var2 = dRows1(0).Item("charValue")
                        '20171117 LEE:
                        var2 = NZ(dRows1(0).Item("charValue"), "[None]")
                    End If
                End If
                contr.Text = var2
                If StrComp(strFld, "3", CompareMethod.Text) = 0 Then
                    'do cbxAcronym also
                    'var2 = dRows1(0).Item("charAcronym")
                    '20171117 LEE:
                    var2 = NZ(dRows1(0).Item("charAcronym"), "[None]")
                    frmH.cbxAssayTechniqueAcronym.Text = var2
                End If
            End If
            If boolCorp Then
                'tbl = tblCorporateAddresses
                tbl = tblCorporateNickNames
                strFld = "charNickname"
                'strFld1 = "id_tblCorporateAddresses"
                strFld1 = "id_tblCorporateNickNames"

                var1 = NZ(drows(0).Item(c.ColumnName), "")
                If Len(var1) = 0 Or var1 = 0 Then
                    var2 = "[None]"
                Else
                    str2 = strFld1 & " = '" & var1 & "'"
                    'dRows1 = tbl.Select(str2, "id_tblCorporateAddresses ASC")
                    dRows1 = tbl.Select(str2, "id_tblCorporateNickNames ASC")
                    If dRows1.Length = 0 Then
                    Else
                        'var2 = dRows1(0).Item(strFld)
                        '20171117 LEE:
                        var2 = NZ(dRows1(0).Item(strFld), "[None]")
                    End If
                End If
                contr.Text = var2
            End If
            If boolEntireR Then
                'Hmmm. Do this in Cancel Report Body Section
            End If

skip1:

        Next


        frmH.dgvStudyConfig.Refresh()

        frmH.dgvDataCompany.AutoResizeColumns()
        frmH.dgvStudyConfig.AutoResizeColumns()


end1:
        boolFromCD = False

        boolStopCBX = False


    End Sub

    Sub ClearGrids()

        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2
        Dim str1 As String
        Dim str2 As String
        Dim varTemp

        'clear data in dgDataWatson
        int1 = tblWatsonData.Rows.Count
        tblWatsonData.Columns.Item(1).ReadOnly = False
        For Count1 = 0 To int1 - 1
            tblWatsonData.Rows.Item(Count1).Item(1) = ""
        Next
        tblWatsonData.Columns.Item(1).ReadOnly = True

        'clear data in dgvDataCompany and dgvStudyConfig
        int1 = tblCompanyData.Rows.Count
        tblCompanyData.Columns.Item(1).ReadOnly = False

        For Count1 = 0 To int1 - 1
            tblCompanyData.Rows.Item(Count1).Item(1) = ""
        Next

        'clear data in dgCompanyAnalRef
        int1 = tblCompanyAnalRefTable.Columns.Count
        For Count1 = int1 - 1 To 0 Step -1
            str1 = tblCompanyAnalRefTable.Columns.Item(Count1).ColumnName
            If StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then
            ElseIf StrComp(str1, "Item", CompareMethod.Text) = 0 Then
            ElseIf StrComp(str1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then
            Else
                var1 = tblCompanyAnalRefTable.Columns.Item(Count1).Caption
                tblCompanyAnalRefTable.Columns.Remove(tblCompanyAnalRefTable.Columns.Item(Count1))
            End If
        Next

        'clear data in dgWatsonAnalRef
        int1 = tblWatsonAnalRefTable.Columns.Count
        For Count1 = int1 - 1 To 1 Step -1
            var1 = tblWatsonAnalRefTable.Columns.Item(Count1).Caption
            'tblWatsonAnalRefTable.Columns.Remove(var1)
            tblWatsonAnalRefTable.Columns.Remove(tblWatsonAnalRefTable.Columns.Item(Count1))
        Next
        frmH.dgvWatsonAnalRef.Refresh()

        'clear data in dgAnalyticalRunSummary
        int1 = tblAnalRunSum.Rows.Count()
        For Count1 = int1 - 1 To 0 Step -1
            tblAnalRunSum.Rows.Remove(tblAnalRunSum.Rows.Item(Count1))
        Next

        'clear data in dgvReports
        'filter tblreports
        'intFilter = dRows(0).Item("id_tblStudies")
        'Dim custTable as System.Data.DataTable = custDS.Tables("Customers")
        'tblReports = tblReports
        'Dim tblReportsView as system.data.dataview = tblReports.DefaultView
        Dim dv As System.Data.DataView = tblReports.DefaultView
        'tblReportsView.Sort = "charReportName"
        str1 = "id_tblStudies = 0" 'intfilter
        dv.RowFilter = str1
        dv.AllowDelete = False
        dv.AllowNew = False
        dv.AllowEdit = False
        frmH.dgvReports.DataSource = dv
        'frmh.dgvReports.Refresh()
        'select first row in dgvReports
        'frmh.dgvReports.Select(0)
        'enter Report Title in lblReportTitle
        frmH.lblReportTitle.Text = ""
        gReportTitle = ""

        'clear data in Summary Table


        'clear data in dgContributing Personnel


        'clear data in Method Validation Data


        'clear data in QA Event Table


        'clear data in Sample Receipt


        'clear data in Appendices


    End Sub
    Sub Prepare_tbl(ByVal tbl As System.Data.DataTable, ByVal dg As DataGrid)

        Dim Count1 As Short
        Dim col1 As New DataColumn
        Dim col2 As New DataColumn

        'format column 1
        col1.DataType = System.Type.GetType("System.String")
        col1.ColumnName = "Item"
        col1.Caption = "Item"
        col1.ReadOnly = True
        tbl.Columns.Add(col1)
        'format column 2
        col2.DataType = System.Type.GetType("System.String")
        col2.ColumnName = "Value"
        col2.Caption = "Value"
        col2.ReadOnly = True
        tbl.Columns.Add(col2)

        'create a new datagridtablestyle
        Dim ts1 As New DataGridTableStyle
        'add the style to the table collection
        ts1.AllowSorting = False
        dg.TableStyles.Add(ts1)

    End Sub

    Sub FindNickname(ByVal cbx As Control, ByVal txt As Control)

        If boolFormLoad Then
            Exit Sub
        End If

        If boolStopCBX Then
            'Exit Sub
        End If

        Dim var1, var2, var3, var4
        Dim Count1 As Short
        Dim Count2 As Short
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim strSQL As String
        Dim ctCols As Short
        Dim ctRows As Short
        Dim strS As String

        'find NickName in da_tblCorporateAddresses
        var1 = cbx.Text

        If StrComp(var1, "[None]", CompareMethod.Text) = 0 Then 'abort
            txt.Text = ""
            GoTo end1
        End If

        tbl = tblCorporateAddresses
        tbl1 = tblCorporateNickNames
        strSQL = "charNickname = '" & var1 & "'"
        Dim drow1() As DataRow = tbl1.Select(strSQL)
        var2 = drow1(0).Item("id_tblCorporateNickNames")

        strSQL = "id_tblCorporateNickNames = " & var2
        strS = "id_tblAddressLabels ASC"
        Dim drow() As DataRow = tbl.Select(strSQL, strS)
        ctRows = drow.Length
        'build Corporate name
        str1 = ""
        ctCols = tbl.Columns.Count
        Count2 = 0
        For Count1 = 0 To ctRows - 1
            var1 = NZ(drow(Count1).Item("charValue"), "")
            var2 = drow(Count1).Item("id_tblAddressLabels")
            var3 = drow(Count1).Item("boolIncludeinTitle")
            var4 = drow(Count1).Item("id_tblAddressLabels")
            If var3 = -1 Then
                str2 = var1 & "*"
            Else
                str2 = var1
            End If
            If StrComp(var1, "[None]", CompareMethod.Text) = 0 Then

            Else
                If Len(NZ(var1, "")) = 0 Then
                Else
                    Count2 = Count2 + 1
                    If Count1 = 0 Then
                        str1 = str2
                    ElseIf var2 = 7 Then 'Or var2 = 7 Then
                        str1 = str1 & ", " & str2
                        'ElseIf var2 = 7 Then
                        '    str1 = str1 & str2
                    ElseIf var2 = 8 Then
                        str1 = str1 & " " & str2
                    Else
                        str1 = str1 & Chr(13) + Chr(10) & str2
                    End If
                End If
            End If
        Next
        txt.Text = str1

end1:

        tbl = Nothing
        drow = Nothing

    End Sub

    Sub ActivateStudyChange()

        'id_tblStudies = 0
        'id_tblReports = 0

        frmH.rbShowIncludedRTConfig.Checked = True
        frmH.rbShowIncludedRBody.Checked = True

        Call GetStudyInfo()

        Call SetDecs()

        Call UpdateTablePropBools()

        'call this to select appropriate value in dgvReportStatements
        Call UpdateWord_dgv()

        Call Set_idtblReports()

        Call MethodValColor()

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        frmH.Refresh()

        boolFormLoad = False

        intOTables = 0



    End Sub

    Sub FillDataCbx()

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim str1 As String
        Dim int1 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int2 As Short
        Dim ctr As ComboBox
        Dim var1, var2

        boolStopCBX = True

        tbl = tblDropdownBoxContent
        For Count1 = 1 To 4
            Select Case Count1
                Case 1 'assay technique
                    ctr = frmH.cbxAssayTechnique
                    int1 = 3
                    str1 = "charValue"
                    strS = "intOrder ASC"
                Case 2 'assay technique acronym
                    ctr = frmH.cbxAssayTechniqueAcronym
                    int1 = 3
                    str1 = "charAcronym"
                    strS = "intOrder ASC"
                Case 3 'anticoagulant
                    ctr = frmH.cbxAnticoagulant
                    int1 = 1
                    str1 = "charValue"
                    strS = "intOrder ASC"
                Case 4 'sample storage temp
                    ctr = frmH.cbxSampleStorageTemp
                    int1 = 9
                    str1 = "charValue"
                    strS = "CHARVALUE ASC"
            End Select
            var2 = ctr.SelectedIndex
            ctr.Items.Clear()
            strF = "id_tblDropdownboxName = " & int1
            Erase rows
            rows = tbl.Select(strF, strS)
            int2 = rows.Length
            For Count2 = 0 To int2 - 1
                var1 = rows(Count2).Item(str1)
                ctr.Items.Add(var1)
            Next
            ctr.SelectedIndex = var2

        Next

        boolStopCBX = False

        'fill cbxxIncSmplDiff
        cbxxIncSmplDiff.Items.Clear()
        cbxxIncSmplDiff.Items.Add("%Difference")
        cbxxIncSmplDiff.Items.Add("Mean %Difference")


    End Sub


    Sub DoThisApplyTemplate()

        Dim boolA As Short
        Dim boolOR As Boolean
        Dim bool As Boolean
        Dim var1


        Call PositionProgress()
        frmH.lblProgress.Text = "Saving configuration..."
        frmH.lblProgress.Visible = True
        frmH.lblProgress.Refresh()

        frmH.panProgress.Visible = True
        frmH.panProgress.Refresh()

        ctPB = 0
        ctPBMax = 12
        ctPB = ctPB + 1
        frmH.pb1.Maximum = ctPBMax
        frmH.pb1.Value = ctPB
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        boolA = -1
        boolOR = True

        'everyone has to SaveHome
        Call SaveHome() 'must be done first in order to retrieve id_tblReports
        Cursor.Current = Cursors.WaitCursor

        '20160711 LEE:
        'If user has permissions to apply template, then user must be able to do the following items
        'so set boolA=-1 for everything


        'boolA = Allowed("boolData") 'rows(0).Item("boolData")
        'ignore BOOLDATA
        boolA = -1
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveDataTabData()
        End If

        'boolA = Allowed("boolSummaryTable") 'rows(0).Item("boolSummaryTable")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            Call SaveSummaryData()
        End If

        'boolA = Allowed("boolAnalRunSummaryTable") 'rows(0).Item("boolAnalRunSummaryTable")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveAnalRunSum()
        End If

        'boolA = Allowed("boolReportTableConfiguration") 'rows(0).Item("boolReportTableConfiguration")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveTableReportData()
        End If


        'boolA =Allowed("boolReportTableHeaderConfig") ' rows(0).Item("boolReportTableHeaderConfig")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveReportTableHeaderConfig()
        End If

        'boolA = Allowed("boolAnalRefStandard") 'rows(0).Item("boolAnalRefStandard")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveCompanyAnalRefTable()
        End If

        'boolA =Allowed("boolContributingPersonnel") ' rows(0).Item("boolContributingPersonnel")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveCP()
        End If

        'boolA = Allowed("boolReportBodySections") 'rows(0).Item("boolReportBodySections")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveReportStatements()
        End If

        'boolA =Allowed("boolMethodValidationData") ' rows(0).Item("boolMethodValidationData")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveMethValTab()
        End If

        'boolA =Allowed("boolQAEventTable") ' rows(0).Item("boolQAEventTable")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            ctPB = ctPB + 1
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()
            Call SaveQATable()
        End If

        'boolA = Allowed("boolSampleReceipt") 'rows(0).Item("boolSampleReceipt")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then
            Call SaveSampleReceiptTab()
        End If

        'boolA =Allowed("boolAppendices") ' rows(0).Item("boolAppendices")
        If boolA = -1 Then
            bool = True
        Else
            bool = False
        End If
        If bool Or boolOR Then

            If boolGuWuOracle Then
                Try
                    ta_tblAppFigs.Update(tblAppFigs)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLAPPFIGS.Merge('ds2005.TBLAPPFIGS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblAppFigsAcc.Update(tblAppFigs)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLAPPFIGS.Merge('ds2005Acc.TBLAPPFIGS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblAppFigsSQLServer.Update(tblAppFigs)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLAPPFIGS.Merge('ds2005Acc.TBLAPPFIGS, True)
                End Try
            End If

        End If

        Call SaveFC()

        'update Summary Table
        Try
            Call UpdateValueSummaryTable()
        Catch ex As Exception

        End Try
        Cursor.Current = Cursors.WaitCursor

        Call LockAll(True, False)

        Call SelectedRefresh()
        Call AssessQCs()

        frmH.pb1.Value = frmH.pb1.Maximum
        frmH.pb1.Refresh()


        Call SetToNonEditMode()

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False

        frmH.panProgress.Visible = False
        frmH.panProgress.Refresh()

        frmH.Refresh()
        'SendKeys.Send("%")
    End Sub

    Sub DoThis(ByVal cmd As String)

        'This routine is for saving frmHome

        Cursor.Current = Cursors.WaitCursor
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short
        Dim strL1 As String
        Dim strL2 As String

        Dim intOT As Short = frmH.lbxTab1.SelectedIndex 'for some reason, after running do this, lbxTab1(n) gets selected if the original index is 6 (Header Config)
        'must put it back at the end

        Dim intRTRT As Short

        'determine index of dgReportTables
        If frmH.dgvReportTables.Rows.Count = 0 Then
            intRTRT = 0
        ElseIf frmH.dgvReportTables.CurrentRow Is Nothing Then
            intRTRT = 0
        Else
            intRTRT = frmH.dgvReportTables.CurrentRow.Index
        End If

        'Call frmH.ShowThis("Beginning")

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow
        rows = tblPermissions.Select(strF)

        'Call frmH.ShowThis("rows = tblPermissions.Select(strF)")

        If StrComp(cmd, "Logoff", CompareMethod.Text) = 0 Then
        Else
            If rows.Length = 0 And boolRefresh = False And StrComp(cmd, "Edit", CompareMethod.Text) = 0 Then
                MsgBox("Guest does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
                Exit Sub
            End If
        End If

        strL1 = "Saving configuration..."

        Cursor.Current = Cursors.WaitCursor

        Dim dgv As DataGridView = frmH.dgvReportTableConfiguration

        Select Case cmd
            Case "Edit"
            Case Else
                dgv.Visible = False
        End Select


        Select Case cmd
            Case "Edit"

                oldCurrentRowRTC = 0
                oldCurrentColRTC = 0
                newCurrentRowRTC = -1
                newCurrentColRTC = -1
                boolRTCEnter = False
                'oldCurrentCellRTC As Object
                'newCurrentCellRTC As Object

                Call LockHomeTab(Not (BOOLHOME))

                frmH.cmdClearStudy.Enabled = False

                Call LockDataTab(Not (BOOLDATA))

                Call LockAnalRunSumTab(Not (BOOLANALRUNSUMMARYTABLE))

                Call LockSummaryTab(Not (BOOLSUMMARYTABLE))


                Call LockReportStatementTab(Not (BOOLREPORTBODYSECTIONS), True)

                'Note:
                'AssignSamples and AdvancedTableConfig goes on in here
                Call LockReportTableTab(Not (BOOLREPORTTABLECONFIGURATION), True, False)

                Call LockReportTableHeaderConfigTab(Not (BOOLREPORTTABLEHEADERCONFIG))

                Call LockAnalRefTab(Not (BOOLANALREFSTANDARD))

                Call LockCPTab(Not (BOOLCONTRIBUTINGPERSONNEL))

                Call LockMethValTab(Not (BOOLMETHODVALIDATIONDATA))

                Call LockQATableTab(Not (BOOLQAEVENTTABLE))

                Call LockSampleReceiptTab(Not (BOOLSAMPLERECEIPT))


                'the rest all called from button
                'App and Figure
                'RW Admin
                'Sample details
                'RW Audit trail

                Call SetToEditMode()

            Case "Save"


                gdtSave = Now

                strL1 = "Saving configuration..."

                Call PositionProgress()
                frmH.lblProgress.Text = strL1 ' "Saving configuration..."
                frmH.lblProgress.Visible = True
                frmH.lblProgress.Refresh()

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

                ctPB = 0
                ctPBMax = 12
                ctPB = ctPB + 1
                frmH.pb1.Maximum = ctPBMax
                frmH.pb1.Value = ctPB
                frmH.pb1.Visible = True
                frmH.pb1.Refresh()


                '*****
                Dim tUserID As String
                Dim tUserName As String

                tUserID = gUserID
                tUserName = gUserName

                strRFC = GetDefaultRFC()
                strMOS = GetDefaultMOS()

                gATAdds = 0
                gATDeletes = 0
                gATMods = 0

                'If gboolAuditTrail And gboolESig Then

                '    Dim frm As New frmESig

                '    frm.ShowDialog()

                '    If frm.boolCancel Then
                '        frm.Dispose()
                '        GoTo end1
                '    End If

                '    gUserID = frm.tUserID
                '    gUserName = frm.tUserName

                '    frm.Dispose()

                'End If

                Dim dt1 As DateTime
                dt1 = Now

                '*****

                'clear audittrailtemp
                tblAuditTrailTemp.Clear()
                idSE = 0

                'Call frmH.ShowThis("Before SaveHome")

                Dim boolDo As Boolean = DoResetFieldCodes()

                boolA = BOOLREPORTTABLECONFIGURATION ' rows(0).Item("boolReportTableConfiguration")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Report Table Configuration Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Dim boolDo1 As Boolean = SaveTableReportData()
                    If boolDo Then
                    Else
                        boolDo = boolDo1
                    End If

                    Call SaveAutoAssignSamples()

                End If


                'everyone has to SaveHome
                Call SaveHome() 'must be done first in order to retrieve id_tblReports
                Cursor.Current = Cursors.WaitCursor

                Call SaveFinalReport()

                strL2 = strL1 & ChrW(10) & "Home Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()

                boolA = BOOLSUMMARYTABLE ' rows(0).Item("boolSummaryTable")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    strL2 = strL1 & ChrW(10) & "Summary Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveSummaryData()
                End If

                'Call frmH.ShowThis("SaveSummaryData")

                boolA = BOOLDATA ' rows(0).Item("boolData")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Data Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveDataTabData()
                    Call SaveFC()
                End If

                'Call frmH.ShowThis("DataTabData")


                boolA = BOOLANALRUNSUMMARYTABLE 'rows(0).Item("boolAnalRunSummaryTable")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Analytical Run Summary Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveAnalRunSum()
                End If

                'Call frmH.ShowThis("SaveAnalRunSum")

                'boolA = BOOLREPORTTABLECONFIGURATION ' rows(0).Item("boolReportTableConfiguration")
                'If boolA = 0 Then
                '    bool = False
                'Else
                '    bool = True
                'End If
                'If bool Or boolOR Then
                '    ctPB = ctPB + 1
                '    frmH.pb1.Value = ctPB
                '    frmH.pb1.Refresh()
                '    strL2 = strL1 & ChrW(10) & "Report Table Configuration Tab..."
                '    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                '    frmH.lblProgress.Refresh()
                '    boolDo = SaveTableReportData()

                '    Call SaveAutoAssignSamples()

                'End If

                'Call frmH.ShowThis("SaveTableReportData")


                boolA = BOOLREPORTTABLEHEADERCONFIG ' rows(0).Item("boolReportTableHeaderConfig")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Report Table Header Config Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveReportTableHeaderConfig()
                End If

                'Call frmH.ShowThis("SaveReportTableHeaderConfig")

                boolA = BOOLANALREFSTANDARD ' rows(0).Item("boolAnalRefStandard")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Analytical Reference Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveCompanyAnalRefTable()
                End If

                'Call frmH.ShowThis("SaveCompanyAnalRefTable")


                boolA = BOOLCONTRIBUTINGPERSONNEL 'rows(0).Item("boolContributingPersonnel")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Contributing Personnel Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveCP()
                End If

                'Call frmH.ShowThis("SaveCP")

                boolA = BOOLREPORTBODYSECTIONS 'rows(0).Item("boolReportBodySections")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Report Template Configuration Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveReportStatements()
                End If

                'Call frmH.ShowThis("SaveReportStatements")

                boolA = BOOLMETHODVALIDATIONDATA 'rows(0).Item("boolMethodValidationData")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "Method Validation Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveMethValTab()
                End If

                'Call frmH.ShowThis("SaveMethValTab")

                boolA = BOOLQAEVENTTABLE ' rows(0).Item("boolQAEventTable")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    ctPB = ctPB + 1
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()
                    strL2 = strL1 & ChrW(10) & "QA Table Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveQATable()
                End If

                'Call frmH.ShowThis("SaveQATable")

                boolA = BOOLSAMPLERECEIPT ' rows(0).Item("boolSampleReceipt")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then
                    strL2 = strL1 & ChrW(10) & "Sample Receipt Tab..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call SaveSampleReceiptTab()
                End If


                'Call frmH.ShowThis("SaveSampleReceiptTab")

                boolA = BOOLAPPENDICES ' rows(0).Item("boolAppendices")
                If boolA = 0 Then
                    bool = False
                Else
                    bool = True
                End If
                If bool Or boolOR Then

                    strL2 = strL1 & ChrW(10) & "Saving Audit Trail Information..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call FillAuditTrailTemp(tblAppFigs)

                    If boolGuWuOracle Then
                        Try
                            ta_tblAppFigs.Update(tblAppFigs)
                        Catch ex As DBConcurrencyException
                            'ds2005.TBLAPPFIGS.Merge('ds2005.TBLAPPFIGS, True)
                        End Try
                    ElseIf boolGuWuAccess Then
                        Try
                            ta_tblAppFigsAcc.Update(tblAppFigs)
                        Catch ex As DBConcurrencyException
                            'ds2005Acc.TBLAPPFIGS.Merge('ds2005Acc.TBLAPPFIGS, True)
                        End Try
                    ElseIf boolGuWuSQLServer Then
                        Try
                            ta_tblAppFigsSQLServer.Update(tblAppFigs)
                        Catch ex As DBConcurrencyException
                            'ds2005Acc.TBLAPPFIGS.Merge('ds2005Acc.TBLAPPFIGS, True)
                        End Try
                    End If

                End If

                'record tblaudittrailtemp
                Call RecordAuditTrail(False, dt1)

                'update Summary Table
                Try
                    Call UpdateValueSummaryTable()
                Catch ex As Exception

                End Try

                'Call frmH.ShowThis("UpdateValueSummaryTable")

                Cursor.Current = Cursors.WaitCursor

                strL2 = strL1 & ChrW(10) & "Final Actions..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()

                Call LockAll(True, False)

                strL2 = strL1 & ChrW(10) & "Selected refresh..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call SelectedRefresh()
                'Call frmH.ShowThis("SelectedRefresh")

                strL2 = strL1 & ChrW(10) & "AssessQCs..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call AssessQCs()
                'Call frmH.ShowThis("AssessQCs")

                If boolDo Then
                    strL2 = strL1 & ChrW(10) & "Reset Field Codes..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call ResetFieldCodes(True)
                End If
                

                'pesky
                strL2 = strL1 & ChrW(10) & "Pesky things..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                If frmH.dgvReportTables.Rows.Count = 0 Then
                Else
                    frmH.dgvReportTables.CurrentCell = frmH.dgvReportTables.Rows(intRTRT).Cells(2)
                    frmH.dgvReportTables.CurrentRow.Selected = True
                    strL2 = strL1 & ChrW(10) & "ReportTableHeaderConfigPopulate..."
                    frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                    frmH.lblProgress.Refresh()
                    Call ReportTableHeaderConfigPopulate()
                End If

                frmH.pb1.Value = frmH.pb1.Maximum
                frmH.pb1.Refresh()


                Call SetToNonEditMode()

                'frmH.lblProgress.Visible = False
                'frmH.pb1.Visible = False

                frmH.panProgress.Visible = False
                frmH.panProgress.Refresh()

                frmH.Refresh()
                'SendKeys.Send("%")

            Case "Cancel"

                boolFormLoad = True

                Cursor.Current = Cursors.WaitCursor

                ctPB = 0
                ctPBMax = 16
                ctPB = ctPB + 1
                frmH.pb1.Maximum = ctPBMax
                frmH.pb1.Value = ctPB
                frmH.pb1.Visible = True
                frmH.pb1.Refresh()

                strL1 = "Canceling changes..."

                strL2 = strL1 & ChrW(10) & "Data Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoDataCancel(False)
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Custom Field Code Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoFCCancel()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Analytical Run Summary Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoAnalRunSumCancel(False)
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Summary Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoCancelSummaryTab()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Analytical Reference Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoAnalRefCancel()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Contributing Personnel Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call doCPCancel()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Method Validation Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call doMethValCancel()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Home Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoHomeCancel() 'do this one last
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Report Template Configuration Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoReportStatementsCancel()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Report Table Configuration Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoRTConfigCancel()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                'strL2 = strL1 & ChrW(10) & "Report Table Header Configuration Tab..."
                strL2 = strL1 & ChrW(10) & "Configure Column Headings Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoCancelRTHConfig()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "QC Table Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoCancelQATable()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                'Call DoCancelAppendix()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()
                strL2 = strL1 & ChrW(10) & "Sample Receipt Tab..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()
                Call DoCancelSampleReceipt()
                ctPB = ctPB + 1
                frmH.pb1.Value = ctPB
                frmH.pb1.Refresh()

                strL2 = strL1 & ChrW(10) & "Final Actions..."
                frmH.lblProgress.Text = strL2 ' "Saving configuration..."
                frmH.lblProgress.Refresh()

                Call LockAll(True, False)

                frmH.pb1.Value = frmH.pb1.Maximum
                frmH.pb1.Refresh()


                Call SetToNonEditMode()


                'pesky

                boolFormLoad = False
                Try
                    Call ReportTableHeaderFilter()
                    Call ReportTableHeaderConfigPopulate()
                Catch ex As Exception


                End Try

                If frmH.dgvReportTables.Rows.Count = 0 Then
                Else
                    frmH.dgvReportTables.CurrentCell = frmH.dgvReportTables.Rows(intRTRT).Cells(2)
                    frmH.dgvReportTables.CurrentRow.Selected = True
                    Call ReportTableHeaderConfigPopulate()
                End If

                Call RTFilter()
                Call DoHomeCancel() 'do this one last

                Call FillFCRW()

                Call FillAnalRunSum()


                'boolFormLoad = False

            Case "cmdExit"

                'Call LockHomeTab(True)
                'Call LockDataTab(True)
                'Call LockSummaryTab(True)
                'Call LockAnalRunSumTab(True)
                'Call LockCPTab(True)
                'Call LockReportTableTab(True)
                'Call LockAnalRefTab(True)
                'Call LockMethValTab(True)
                'Call LockReportStatementTab(True)
                'Call LockReportTableHeaderConfigTab(True)
                'Call LockQATableTab(True)
                'Call LockSampleReceiptTab(True)
                ''Call LockAppendixTab(True)
                'Call LockAdministration(True)

                Call LockAll(True, False)

                Call SetToNonEditMode()

        End Select

        '20190220 LEE: Pesky
        Call ColorMethodValRows()

        'frmH.pb1.Visible = False


        'do some stuff in HomeTab one more time
        frmH.optStudyDocStudies.Enabled = True
        frmH.optStudyDocOpen.Enabled = True
        frmH.optStudyDocClosed.Enabled = True

        'MsgBox("1")

        If frmH.cmdEdit.Enabled Then
            If frmH.rbOracle.Checked Then
                frmH.panFilterStudy.Enabled = True
                frmH.panStudyFilter.Enabled = True
                frmH.gbStudyFilter.Enabled = True
            Else
                frmH.panFilterStudy.Enabled = False
                frmH.panStudyFilter.Enabled = False
                frmH.gbStudyFilter.Enabled = False
            End If
        Else
            frmH.panFilterStudy.Enabled = False
            frmH.panStudyFilter.Enabled = False
            frmH.gbStudyFilter.Enabled = False
        End If

        If frmH.rbOracle.Checked And boolFormLoad Then
            frmH.panFilterStudy.Enabled = True
            frmH.panStudyFilter.Enabled = True
            frmH.gbStudyFilter.Enabled = True

            'pesky
            Call SetPanAction()

        End If

        ' MsgBox("2")

        Select Case intOT
            Case 6, 12, 13, 14, 15
                GoTo end1
        End Select


        Cursor.Current = Cursors.Default

        dgv.Visible = True


        frmH.Refresh()

        'MsgBox("3")

        'make sure cbxStudies isn't selected
        frmH.lbxTab1.Select()

        'pesky
        Call frmH.ViewSections(Not (boolEntireReport))
        'MsgBox("4")
        Try
            frmH.dgvMethodValData.AutoResizeRows()

        Catch ex As Exception

        End Try

        Call SetComboCell(frmH.dgvReportTableConfiguration, "CHARPAGEORIENTATION")

        'MsgBox("5")

        'Try
        '    'keep in Try
        '    'tblPermissions may not be open yet
        '    boolA = Allowed("BOOLANALRUNSUMMARYTABLE")
        '    If boolA Then
        '        Try
        '            frmH.cmdViewAnalyticalRuns1.Enabled = True
        '        Catch ex As Exception

        '        End Try

        '    End If
        'Catch ex As Exception

        'End Try


end1:
        'Call frmH.ShowThis("frmH.ViewSections(Not (boolEntireReport)")

        dgv.Visible = True

        'MsgBox("6")

    End Sub

    Sub LockSummaryTab(ByVal bool)

        frmH.panSumTable.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp4.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 _
                     Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0 _
                     Or InStr(1, ctrl.Name, "grpMethodValidation", CompareMethod.Text) > 0)) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception
                End Try
            End If
        Next

        frmH.dgvSummaryData.Enabled = True


        Dim boolA As Boolean
        boolA = BOOLSUMMARYTABLE

        If boolA Then
            Dim int1 As Short
            Dim Count1 As Short

            frmH.dgvSummaryData.ReadOnly = bool
            'int1 = frmh.dgvSummaryData.Columns.Count
            'For Count1 = 0 To int1 - 1
            '    frmh.dgvSummaryData.Columns.item(Count1).ReadOnly = True
            'Next
            'frmh.dgvSummaryData.Columns.item("boolInclude").ReadOnly = bool
            'frmh.dgvSummaryData.Columns.item("intOrder").ReadOnly = bool

            frmH.dgvSummaryData.Columns.Item("charRowName").ReadOnly = True
            frmH.dgvSummaryData.Columns.Item("charValue").ReadOnly = True


            frmH.cmdResetSummaryTable.Enabled = Not (bool)
            frmH.cmdOrderSummaryTable.Enabled = Not (bool)

        End If

    End Sub


    Sub LockMethValTab(ByVal bool)

        frmH.panMethVal.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp10.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

        frmH.dgvMethodValData.Enabled = True
        frmH.dgvMethValExistingGuWu.Enabled = True



        Dim Count1 As Short
        Dim dgv As DataGridView
        Dim var1, var2

        Try
            dgv = frmH.dgvMethodValData

            dgv.ReadOnly = bool
            Try
                dgv.Columns(0).ReadOnly = True
            Catch ex As Exception

            End Try

            'if Sample Analysis and tblstudies2 <> null, then read-only
            Dim tblM As System.Data.DataTable
            Dim strF As String
            Dim rowsM() As DataRow
            tblM = tblMethodValidationData

            Dim strT As String
            Dim rowsT() As DataRow
            Dim boolSamples As Boolean = False
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            rowsT = tblReports.Select(strF)
            strT = NZ(rowsT(0).Item("CHARREPORTTYPE"), "Sample Analysis")
            If InStr(1, strT, "Sample", CompareMethod.Text) > 0 Then
                boolSamples = True
            Else
                boolSamples = False
            End If
            If boolSamples And bool = False Then
                For Count1 = 1 To dgv.Columns.Count - 1
                    Try
                        var1 = dgv.Columns(Count1).Name
                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARCOLUMNNAME = '" & CleanText(CStr(var1)) & "'"
                        rowsM = tblM.Select(strF)
                        If rowsM.Length = 0 Then
                            dgv.Columns(Count1).ReadOnly = False
                        Else
                            var2 = rowsM(0).Item("ID_TBLSTUDIES2")
                            If IsDBNull(var2) Or var2 = 0 Then
                                dgv.Columns(Count1).ReadOnly = bool
                            Else
                                dgv.Columns(Count1).ReadOnly = True
                            End If
                        End If
                    Catch ex As Exception

                    End Try

                Next
            Else
                For Count1 = 1 To dgv.Columns.Count - 1
                    dgv.Columns(Count1).ReadOnly = bool
                Next
            End If

            'If InStr(1, strT, "Sample", CompareMethod.Text) > 0 Then
            '    For Count1 = 1 To dgv.Columns.Count - 1
            '        dgv.Columns(Count1).ReadOnly = True
            '    Next
            'Else
            '    For Count1 = 1 To dgv.Columns.Count - 1
            '        dgv.Columns(Count1).ReadOnly = bool
            '    Next
            'End If

            'frmh.dgMethValExistingGuWu.ReadOnly = bool

            frmH.gbMethodValMultiple.Enabled = Not (bool)
            'frmh.gbMethValApplyGuWu.Enabled = Not (bool)

            frmH.cmdMethValExecute.Enabled = Not (bool)
            frmH.cmdMethValReset.Enabled = Not (bool)

            frmH.cmdBrowseMDB.Enabled = Not (bool)
            frmH.cmdMethValUpdate.Enabled = Not (bool)

            frmH.cbxArchivedMDB.Enabled = Not (bool)
            frmH.cbxMethValExistingGuWu.Enabled = Not (bool)


        Catch ex As Exception

        End Try



    End Sub

    Sub LockQATableTab(ByVal bool)

        frmH.panQAEvent.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp11.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception
                End Try
            End If
        Next

        frmH.dgQATable.Enabled = True


        frmH.cmdInsertQAEvent.Enabled = Not (bool)
        frmH.cmdInsertQAEvent.Enabled = Not (bool)
        frmH.cmdDeleteQAEvent.Enabled = Not (bool)
        frmH.cmdQACancel.Enabled = Not (bool)

        frmH.dgQATable.ReadOnly = bool

        frmH.chkQAEventBorder.Enabled = Not (bool)



    End Sub

    Sub LockAppendixTab(bool As Boolean)

        Dim ctrl As Control
        For Each ctrl In frmH.tp13.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

    End Sub

    Sub LockSampleDetails(bool As Boolean)

        Dim ctrl As Control
        For Each ctrl In frmH.tp15.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then

                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

    End Sub

    Sub LockRWAuditTrail(bool As Boolean)

        Dim ctrl As Control
        For Each ctrl In frmH.tp16.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then

                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next
        'button is opposite of edit
        Dim boolA As Boolean = BOOLRWAUDITTRAIL
        If boolA Then
            frmH.cmdAuditTrail.Enabled = bool
        End If

    End Sub

    Sub LockReportTableTab(bool As Boolean, boolDoA As Boolean, boolFromAll As Boolean)

        'booldoA means: do AssignSamples and Advanced Table Config

        Dim var1

        'Note: has asynchronous buttons in panRepTables
        frmH.panRepTables.Enabled = True

        'evaluate only labels in this tab
        Dim ctrl As Control
        For Each ctrl In frmH.tp6.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 _
                    Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0 _
                    Or InStr(1, ctrl.Name, "Filter", CompareMethod.Text) > 0 _
                    Or InStr(1, ctrl.Name, "TableGraphicExamples", CompareMethod.Text) > 0 _
                    )) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

        For Each ctrl In frmH.panRepTables.Controls
            Try
                ctrl.Enabled = Not (bool)
            Catch ex As Exception

            End Try
        Next


        Dim boolA As Boolean
        boolA = BOOLREPORTTABLECONFIGURATION

        If boolA Then
        Else
            bool = True
        End If

        For Each ctrl In frmH.panRepTables.Controls
            Try
                Select Case ctrl.Name
                    Case "cmdCreateTable"
                        ctrl.Enabled = bool
                    Case Else
                        ctrl.Enabled = Not (bool)
                End Select

            Catch ex As Exception

            End Try
        Next

        frmH.gbRTC.Enabled = True
        frmH.gbFilters.Enabled = True

        'cmdAssignSample is opposite
        frmH.cmdAssignSamples.Enabled = bool

        Dim gs As DataGridColumnStyle
        Dim Count1 As Short
        Dim int1 As Short

        frmH.cmdRTConfigCancel.Enabled = Not (bool)
        frmH.chkReadOnlyTables.Enabled = Not (bool)

        frmH.cmdOrderReportTableConfig.Enabled = Not (bool)
        frmH.cmdImportTables.Enabled = Not (bool)

        frmH.cmdRTCDown.Enabled = Not (bool)
        frmH.cmdRTCUp.Enabled = Not (bool)

        frmH.cmdUpA.Enabled = Not (bool)
        frmH.cmdDownA.Enabled = Not (bool)

        frmH.cmdUpCF.Enabled = Not (bool)
        frmH.cmdDownCF.Enabled = Not (bool)

        frmH.dgvReportTableConfiguration.ReadOnly = bool
        int1 = frmH.dgvReportTableConfiguration.Columns.Count
        frmH.dgvReportTableConfiguration.ReadOnly = bool
        frmH.dgvReportTableConfiguration.Columns.Item("CHARTABLENAME").ReadOnly = True
        frmH.dgvReportTableConfiguration.Columns.Item("CHARHEADINGTEXT").ReadOnly = False 'True 20170522 LEE: Allow editing here
        'frmh.dgvReportTableConfiguration.Columns.item("boolRequiresSampleAssignment").ReadOnly = True

        'Note: cmdEdit HASN'T been changed yet
        boolA = BOOLADVANCEDTABLE
        If boolA Then
            'If frmH.cmdEdit.Enabled Then
            '    frmH.cmdAdvancedTable.Enabled = True
            'Else
            '    frmH.cmdAdvancedTable.Enabled = False ' Not (bool)
            'End If
            frmH.cmdAdvancedTable.Enabled = True
        Else
            frmH.cmdAdvancedTable.Enabled = False
        End If

        'these two always stay enabled
        frmH.cmdAssignSamples.Enabled = True
        frmH.cmdOutliers.Enabled = True
        frmH.cmdViewAnalRuns.Enabled = True
        frmH.chkTableName.Enabled = True
        frmH.chkTableGraphicExamples.Enabled = True
        frmH.cmdShowGroups.Enabled = True
        frmH.dgvGroups.Enabled = True
        frmH.cmdResize.Enabled = True


    End Sub


    Sub LockAll(bool As Boolean, boolFromAdmin As Boolean)

        Call LockHomeTab(bool)
        Call LockDataTab(bool)
        Call LockSummaryTab(bool)
        Call LockAnalRunSumTab(bool)
        Call LockCPTab(bool)
        Call LockReportTableTab(bool, Not (bool), boolFromAdmin)
        Call LockAnalRefTab(bool)
        Call LockMethValTab(bool)
        Call LockReportStatementTab(bool, Not (bool))
        Call LockReportTableHeaderConfigTab(bool)
        Call LockQATableTab(bool)
        Call LockSampleReceiptTab(bool)
        'Call LockAppendixTab(bool)
        'Call LockAdministration(bool)
        'Call LockSampleDetails(bool)
        'Call LockRWAuditTrail(bool)

    End Sub

    Sub DoAnalRunSumCancel(ByVal boolFromReset As Boolean)
        Call FillReportTableHome()
        Call FillAnalRunSum()
    End Sub

    Sub SaveFC()

        '2218
        Try
            Call FillAuditTrailTemp(tblCustomFieldCodes)

            If boolGuWuOracle Then
                Try
                    ta_tblAuditTrail.Update(tblAuditTrail)
                Catch ex As DBConcurrencyException
                    'ds2005.tblAuditTrail.Merge('ds2005.tblAuditTrail, True)
                End Try

            ElseIf boolGuWuAccess Then
                Try
                    ta_tblCustomFieldCodesAcc.Update(tblCustomFieldCodes)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLCUSTOMFIELDCODES.Merge('ds2005Acc.TBLCUSTOMFIELDCODES, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblCustomFieldCodesSQLServer.Update(tblCustomFieldCodes)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLCUSTOMFIELDCODES.Merge('ds2005Acc.TBLCUSTOMFIELDCODES, True)
                End Try
            End If
        Catch ex As Exception

        End Try


    End Sub

    Sub SaveAnalRunSum()

        Dim tbl As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim strF As String
        Dim str1 As String
        Dim str2 As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Int16
        Dim var1, var2, var3, var4, var5
        Dim dv As System.Data.DataView
        Dim drv As DataRowView
        Dim drows() As DataRow
        Dim boolI As Boolean
        Dim boolIRegr As Boolean
        Dim intUL As Short = 1

        dv = frmH.dgvAnalyticalRunSummary.DataSource
        tbl = dv.ToTable()


        'tbl = tblAnalRunSum
        dtbl = tblAnalyticalRunSummary
        'dv = dtbl.DefaultView
        'dv.AllowNew = True
        'dv.AllowEdit = True

        'loop through tbl, find appropriate dv entries, then record comment value in dtbl
        int1 = tbl.Rows.Count
        For Count1 = 0 To int1 - 1
            var1 = NZ(tbl.Rows.Item(Count1).Item("Watson Run ID"), "") 'Watson Run ID
            int2 = Len(var1)
            If Len(var1) = 0 Then
            Else
                'find appropriate entry in dv
                var4 = NZ(tbl.Rows.Item(Count1).Item("Analyte"), "") 'Analyte
                var2 = NZ(tbl.Rows.Item(Count1).Item("Analyte_C"), "") 'Analyte
                strF = "id_tblStudies = " & id_tblStudies & " and intWatsonRunID = " & var1 & " and charAnalyte = '" & CleanText(CStr(var2)) & "'"

                '20170927 LEE: If study is single analyte, multiple matrix, the previous filter isn't returning appropiate records
                'must loop and do both Analyte and Analyte_C

                If StrComp(var2, var4, CompareMethod.Text) = 0 Then
                    intUL = 1
                Else
                    intUL = 2
                End If

                For Count3 = 1 To intUL

                    Select Case Count3
                        Case 1
                            var5 = var2
                        Case 2
                            var5 = var4
                    End Select

                    strF = "id_tblStudies = " & id_tblStudies & " and intWatsonRunID = " & var1 & " and charAnalyte = '" & CleanText(CStr(var5)) & "'"

                    drows = dtbl.Select(strF)
                    int3 = drows.Length
                    'dv.RowFilter = strF
                    var3 = NZ(tbl.Rows.Item(Count1).Item("User Comments"), "")
                    boolI = tbl.Rows.Item(Count1).Item("boolInclude")
                    boolIRegr = NZ(tbl.Rows.Item(Count1).Item("boolIncludeRegr"), True)
                    'If dv.Count = 0 Then 'add new record
                    If drows.Length = 0 Then 'add new record
                        Dim drow As DataRow = dtbl.NewRow()
                        drow.BeginEdit()
                        drow.Item("charUserComments") = var3
                        drow.Item("id_tblStudies") = id_tblStudies
                        drow.Item("intWatsonRunID") = var1
                        'drow.Item("charAnalyte") = var2
                        drow.Item("charAnalyte") = var5
                        drow.Item("boolInclude") = boolI
                        drow.Item("boolIncludeRegr") = boolIRegr
                        drow.EndEdit()
                        dtbl.Rows.Add(drow)
                    Else 'modify existing record
                        'record entry
                        drows(0).BeginEdit()
                        drows(0).Item("charUserComments") = var3
                        drows(0).Item("boolInclude") = boolI
                        drows(0).Item("boolIncludeRegr") = boolIRegr
                        drows(0).EndEdit()
                    End If

                Next

                'drows = dtbl.Select(strF)
                'int3 = drows.Length
                ''dv.RowFilter = strF
                'var3 = NZ(tbl.Rows.Item(Count1).Item("User Comments"), "")
                'boolI = tbl.Rows.Item(Count1).Item("boolInclude")
                'boolIRegr = NZ(tbl.Rows.Item(Count1).Item("boolIncludeRegr"), True)
                ''If dv.Count = 0 Then 'add new record
                'If drows.Length = 0 Then 'add new record
                '    Dim drow As DataRow = dtbl.NewRow()
                '    drow.BeginEdit()
                '    drow.Item("charUserComments") = var3
                '    drow.Item("id_tblStudies") = id_tblStudies
                '    drow.Item("intWatsonRunID") = var1
                '    drow.Item("charAnalyte") = var2
                '    drow.Item("boolInclude") = boolI
                '    drow.Item("boolIncludeRegr") = boolIRegr
                '    drow.EndEdit()
                '    dtbl.Rows.Add(drow)
                'Else 'modify existing record
                '    'record entry
                '    drows(0).BeginEdit()
                '    drows(0).Item("charUserComments") = var3
                '    drows(0).Item("boolInclude") = boolI
                '    drows(0).Item("boolIncludeRegr") = boolIRegr
                '    drows(0).EndEdit()
                'End If

             
            End If
        Next


        'Dim dvCheck as system.data.dataview = New DataView(tblAnalyticalRunSummary)
        'dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblAnalyticalRunSummary)

            If boolGuWuOracle Then
                Try
                    ta_tblAnalyticalRunSummary.Update(tblAnalyticalRunSummary)
                Catch ex As DBConcurrencyException
                    Try
                        'ds2005.TBLANALYTICALRUNSUMMARY.Merge('ds2005.TBLANALYTICALRUNSUMMARY, True)

                    Catch ex1 As Exception
                        MsgBox("TBLANALYTICALRUNSUMMARY: " & ex1.Message)
                    End Try
                    ''ds2005.TBLANALYTICALRUNSUMMARY.Merge('ds2005.TBLANALYTICALRUNSUMMARY, True)

                End Try

                'sometimes this table will not update. Try running it twice.
                Try
                    ta_tblAnalyticalRunSummary.Update(tblAnalyticalRunSummary)
                Catch ex As DBConcurrencyException
                    Try
                        'ds2005.TBLANALYTICALRUNSUMMARY.Merge('ds2005.TBLANALYTICALRUNSUMMARY, True)

                    Catch ex1 As Exception
                        MsgBox("TBLANALYTICALRUNSUMMARY: " & ex1.Message)
                    End Try
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblAnalyticalRunSummaryAcc.Update(tblAnalyticalRunSummary)
                Catch ex As DBConcurrencyException

                    'ds2005Acc.TBLANALYTICALRUNSUMMARY.Merge('ds2005Acc.TBLANALYTICALRUNSUMMARY, True)

                End Try

                'sometimes this table will not update. Try running it twice.
                Try
                    ta_tblAnalyticalRunSummaryAcc.Update(tblAnalyticalRunSummary)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLANALYTICALRUNSUMMARY.Merge('ds2005Acc.TBLANALYTICALRUNSUMMARY, True)
                End Try

            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblAnalyticalRunSummarySQLServer.Update(tblAnalyticalRunSummary)
                Catch ex As DBConcurrencyException

                    'ds2005Acc.TBLANALYTICALRUNSUMMARY.Merge('ds2005Acc.TBLANALYTICALRUNSUMMARY, True)

                End Try

                'sometimes this table will not update. Try running it twice.
                Try
                    ta_tblAnalyticalRunSummarySQLServer.Update(tblAnalyticalRunSummary)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLANALYTICALRUNSUMMARY.Merge('ds2005Acc.TBLANALYTICALRUNSUMMARY, True)
                End Try

            End If

        End If



    End Sub


    Sub LockAnalRunSumTab(ByVal bool As Boolean)

        Dim var1
        Dim ctrl As Control
        For Each ctrl In frmH.tp3.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 _
                    Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0 _
                    Or InStr(1, ctrl.Name, "grpReviewAnalyticalRuns", CompareMethod.Text) > 0)) Then
                Try
                    ctrl.Enabled = False 'Not (bool)
                Catch ex As Exception
                End Try
            End If
        Next

        'leave dgv enabled
        frmH.dgvAnalyticalRunSummary.Enabled = True

        For Each ctrl In frmH.panAnalRuns.Controls
            Try
                ctrl.Enabled = False ' Not (bool)
            Catch ex As Exception

            End Try
        Next

        Dim boolA As Boolean = BOOLANALRUNSUMMARYTABLE

        If boolA Then

            For Each ctrl In frmH.panAnalRuns.Controls
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            Next

            frmH.cmdViewAnalyticalRuns1.Enabled = True

            Dim int1 As Short
            Dim Count1 As Short

            Dim dgv As DataGridView = frmH.dgvAnalyticalRunSummary

            'frmh.dganalyticalRunSummary.TableStyles(0).GridColumnStyles(8).ReadOnly = bool
            'frmh.dganalyticalRunSummary.ReadOnly = bool
            'frmh.dganalyticalRunSummary.Refresh()
            frmH.panAnalRunSum.Enabled = Not (bool)
            frmH.panAnalRunChoices.Enabled = Not (bool)
            frmH.cmdAnaRunSumCancel.Enabled = Not (bool)
            frmH.gbReportOptions.Enabled = Not (bool)

            dgv.ReadOnly = bool
            int1 = dgv.Columns.Count
            For Count1 = 0 To int1 - 1
                dgv.Columns.Item(Count1).ReadOnly = True
            Next
            dgv.Columns.Item("boolInclude").ReadOnly = bool
            dgv.Columns.Item("boolIncludeRegr").ReadOnly = bool
            dgv.Columns.Item("User Comments").ReadOnly = bool

            dgv.Refresh()
        Else

        End If

        var1 = "a" 'debug

    End Sub

    Sub DoAnalRefCancel()

        Dim dg As DataGrid
        Dim int1 As Short
        Dim tbl As System.Data.DataTable

        tbl = tblCompanyAnalRefTable
        int1 = tbl.Columns.Count

        Call AddColumnsAnalRefTable()
        Call FillCompanyAnalRefTable()
        Call ResizeDV(frmH.dgvCompanyAnalRef, False)

        frmH.dgvCompanyAnalRef.AutoResizeColumns()
        frmH.dgvWatsonAnalRef.AutoResizeColumns()

        frmH.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
        Call SyncCols(frmH.dgvWatsonAnalRef, frmH.dgvCompanyAnalRef)


    End Sub

    Sub DoDataCancel(ByVal boolFromReset As Boolean)

        Dim var1
        Dim strF As String
        Dim strS As String

        '*****

        'do AnalyteSort Tab
        tblAnalytesHome.RejectChanges()
        tblAnalyteGroups.RejectChanges()

        'now must reset tblanalyteids and tblmatrices
        strF = "INTGROUP > -2"
        strS = "INTORDER ASC"
        Dim dvAnalyteGroups As New DataView(tblAnalyteGroups, strF, strS, DataViewRowState.CurrentRows)
        tblAnalyteIDs = dvAnalyteGroups.ToTable("tblAnalyteIDs", True, "ANALYTEID", "ANALYTEDESCRIPTION")

        'enter data for Table Matrices
        tblMatrices = dvAnalyteGroups.ToTable("tblMatrices", True, "MATRIX")

        var1 = tblMatrices.Rows.Count
        var1 = var1

        Call ReorderAnalytes()

        'now reorder Report Tables
        Cursor.Current = Cursors.WaitCursor
        Call FillTableReports(True)
        Cursor.Current = Cursors.WaitCursor
        Call FillTableReportsAnalytes(True)
        Cursor.Current = Cursors.WaitCursor
        Call FillTableReportDataAnalytes(True)
        Call RTFilter()
        Cursor.Current = Cursors.WaitCursor

        '*****

        Call FillDataTabData(boolFromReset)

    End Sub

    Sub DoFCCancel()

        tblCustomFieldCodes.RejectChanges()

    End Sub

    Sub DoHomeCancel()

        Dim dv As System.Data.DataView
        Dim strFilter As String
        Dim int1 As Short
        Dim Count1 As Short

        strFilter = "id_tblStudies = " & id_tblStudies

        'first delete any newly-added rows

        tblReports.RejectChanges()

        'dv = New DataView(tblReports, strFilter, "id_tblReports", DataViewRowState.Added)
        'int1 = dv.Count
        'For Count1 = int1 - 1 To 0 Step -1
        '    dv.Delete(Count1)
        'Next

        'now delete contents of cbxxReport
        'cbxxReportTypes.Value = ""
        'cbxxReportTemplates.Value = ""

        dv = New DataView(tblReports, strFilter, "id_tblReports", DataViewRowState.OriginalRows)
        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = True
        frmH.dgvReports.DataSource = dv
        frmH.dgvReports.Refresh()
        'End If

        'out
        Call TempReportsOrder() 'for some reason, column order of last two columns gets switched

        'out
        Call SetReportConfigType()

        'pesky
        'Call FillHomeDropdownBoxes()

        'Call SetReportTableColumns()

        frmH.dgvReports.AutoResizeRows()



    End Sub

    Sub SetReportTableColumns()

        Exit Sub 'don't do this anymore

        Dim int1 As Short
        Dim int2 As Short
        Dim dgv As DataGridView
        Dim str1 As String

        dgv = frmH.dgvReports
        If dgv.RowCount = 0 Then
            id_tblConfigReportType = -1
            Exit Sub
        Else
            If dgv.CurrentRow Is Nothing Then
                int1 = 0
            Else
                int1 = dgv.CurrentRow.Index
            End If
        End If

        str1 = NZ(dgv.Item("CHARREPORTTYPE", int1).Value, "Sample Analysis")
        cbxxReportTypes.Value = str1


    End Sub

    Sub SaveCP()

        Dim dv As System.Data.DataView
        Dim tbl As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim strF As String
        Dim ct1 As Short
        Dim ct2 As Short
        Dim ct3 As Short
        Dim drows1() As DataRow
        Dim dv2 As System.Data.DataView
        Dim row As DataRow
        Dim dr() As DataRow
        Dim boolExists As Boolean
        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2, var3
        Dim int1 As Short
        Dim col As DataColumn
        Dim str1 As String
        Dim dv1 As System.Data.DataView

        frmH.dgvContributingPersonnel.CommitEdit(DataGridViewDataErrorContexts.Commit)

        tbl1 = tblContributingPersonnel
        ct1 = tbl1.Rows.Count
        dv = frmH.dgvContributingPersonnel.DataSource
        ct2 = dv.Count

        'determine if criteria is met
        'must have a name
        dv.AllowDelete = True
        For Count1 = ct2 - 1 To 0 Step -1
            If dv(Count1).Row.RowState = DataRowState.Deleted Then 'ignore
            Else
                var1 = NZ(dv(Count1).Item("charCPName"), "")
                If Len(var1) = 0 Then 'delete
                    dv(Count1).Row.Delete()
                End If
            End If
        Next
        dv.AllowDelete = False

        'now fix intOrder values
        str1 = "intOrder ASC"
        dv.Sort = str1
        ct2 = dv.Count

        For Count1 = 0 To ct2 - 2
            var1 = NZ(dv(Count1).Item("intorder"), 0)
            If Count1 = 0 And var1 = 0 Then
                var1 = 1
                dv(Count1).Row.BeginEdit()
                dv(Count1).Item("intOrder") = var1
                dv(Count1).Row.EndEdit()
            End If
            var2 = NZ(dv(Count1 + 1).Item("intorder"), 0)
            If var1 + 1 = var2 Then 'proceed
            Else 'fix
                var2 = var1 + 1
                dv(Count1).Row.BeginEdit()
                dv(Count1 + 1).Item("intOrder") = var2
                dv(Count1 + 1).Row.EndEdit()
            End If
        Next

        'for some reason, BOOLINCLUDESIGONTABLEPAGE is null instead of zero
        For Count1 = 0 To ct2 - 1
            var1 = dv(Count1).Item("BOOLINCLUDESIGONTABLEPAGE")
            If IsDBNull(var1) Then
                dv(Count1).BeginEdit()
                dv(Count1).Item("BOOLINCLUDESIGONTABLEPAGE") = 0
                dv(Count1).EndEdit()
            End If
        Next

        Dim dvCheck As System.Data.DataView = New DataView(tblContributingPersonnel)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblContributingPersonnel)

            If boolGuWuOracle Then
                Try
                    ta_tblContributingPersonnel.Update(tblContributingPersonnel)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005.TBLCONTRIBUTINGPERSONNEL, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblContributingPersonnelAcc.Update(tblContributingPersonnel)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005Acc.TBLCONTRIBUTINGPERSONNEL, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblContributingPersonnelSQLServer.Update(tblContributingPersonnel)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLCONTRIBUTINGPERSONNEL.Merge('ds2005Acc.TBLCONTRIBUTINGPERSONNEL, True)
                End Try
            End If

        End If


    End Sub
    Sub SaveHome()

        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim dtbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim bool As Boolean
        Dim dgv As DataGridView
        Dim rows() As DataRow
        Dim Count1 As Short

        dgv = frmH.dgvReports

        Call FillAuditTrailTemp(tblReports)

        'record appropriate values in tblReports
        str1 = "id_tblStudies = " & id_tblStudies
        dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        dtbl = tblReports

        rows = dtbl.Select(str1)
        If rows.Length = 0 Then 'first save the record then continue

            If boolGuWuOracle Then
                Try
                    ta_tblReports.Update(tblReports)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLREPORTS.Merge('ds2005.TBLREPORTS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportsAcc.Update(tblReports)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTS.Merge('ds2005Acc.TBLREPORTS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportsSQLServer.Update(tblReports)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTS.Merge('ds2005Acc.TBLREPORTS, True)
                End Try
            End If

            dtbl = tblReports
            rows = dtbl.Select(str1)

        End If

        dv = dgv.DataSource
        If dv.Count = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow.Index = Nothing Then
            int1 = 0
        Else
            int1 = dgv.CurrentRow.Index
        End If

        'dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        rows(int1).BeginEdit()

        'record intCalStd 
        int2 = 1 'retired field
        rows(int1).Item("intCalStd") = int2

        'record intQC
        int2 = 1 'retired field
        rows(int1).Item("intQC") = int2

        'record intShowBQL
        int2 = 1 'retired field
        rows(int1).Item("intShowBQL") = int2

        'record intShowCalStd
        int2 = 1 'retired field
        rows(int1).Item("intShowCalStd") = int2

        'record intUserComments
        int2 = 1
        If frmH.rbUseWatsonComments.Checked Then
            int2 = 1
        ElseIf frmH.rbUseUserComments.Checked Then
            int2 = 2
        End If
        rows(int1).Item("intUserComments") = int2

        int2 = 1 'retired field
        rows(int1).Item("BOOLEXCLUDEPSAE") = int2

        'BOOLALLAR chkAll
        'BOOLACCAR chkAccepted
        'BOOLREJAR chkRejected
        'BOOLREGRAR chkRegrPerformed
        'BOOLNOREGRAR chkNoRegrPerformed
        'BOOLINCLPSAE chkPSAE

        If frmH.chkAll.Checked Then

        End If
        Dim chk As CheckBox
        For count1 = 1 To 6
            Select Case Count1
                Case 1
                    chk = frmH.chkAll
                    str1 = "BOOLALLAR"
                Case 2
                    chk = frmH.chkAccepted
                    str1 = "BOOLACCAR"
                Case 3
                    chk = frmH.chkRejected
                    str1 = "BOOLREJAR"
                Case 4
                    chk = frmH.chkRegrPerformed
                    str1 = "BOOLREGRAR"
                Case 5
                    chk = frmH.chkNoRegrPerformed
                    str1 = "BOOLNOREGRAR"
                Case 6
                    chk = frmH.chkPSAE
                    str1 = "BOOLINCLPSAE"
            End Select

            If chk.Checked Then
                int2 = -1
            Else
                int2 = 0
            End If
            rows(int1).Item(str1) = int2

        Next

        If gboolDisplayAttachments Then
            rows(int1).Item("BOOLDISPLAYATTACHMENTS") = -1
        Else
            rows(int1).Item("BOOLDISPLAYATTACHMENTS") = 0
        End If

        If gboolReadOnlyTables Then
            rows(int1).Item("BOOLREADONLYTABLES") = -1
        Else
            rows(int1).Item("BOOLREADONLYTABLES") = 0
        End If

        'record id_tblStudies
        rows(int1).Item("id_tblStudies") = id_tblStudies

        rows(int1).EndEdit()


        Dim dvCheck As System.Data.DataView = New DataView(tblReports)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblReports)

            If boolGuWuOracle Then
                Try
                    ta_tblReports.Update(tblReports)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaHome: " & ex.Message)
                    'ds2005.TBLREPORTS.Merge('ds2005.TBLREPORTS, True)
                End Try

                Try
                    ta_tblReportHeaders.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLREPORTHEADERS.Merge('ds2005.TBLREPORTHEADERS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportsAcc.Update(tblReports)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaHome: " & ex.Message)
                    'ds2005Acc.TBLREPORTS.Merge('ds2005Acc.TBLREPORTS, True)
                End Try

                Try
                    ta_tblReportHeadersAcc.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTHEADERS.Merge('ds2005Acc.TBLREPORTHEADERS, True)
                End Try

            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportsSQLServer.Update(tblReports)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaHome: " & ex.Message)
                    'ds2005Acc.TBLREPORTS.Merge('ds2005Acc.TBLREPORTS, True)
                End Try

                Try
                    ta_tblReportHeadersSQLServer.Update(tblReportHeaders)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTHEADERS.Merge('ds2005Acc.TBLREPORTHEADERS, True)
                End Try

            End If

        End If

        str1 = NZ(rows(int1).Item("charReportTitle"), "")
        frmH.lblReportTitle.Text = str1
        gReportTitle = str1

        ''remove eventhandler
        'RemoveHandler combo.SelectedIndexChanged, New EventHandler(AddressOf ComboBox_SelectedIndexChanged)
        'RemoveHandler Obj.Ev_Event, AddressOf EventHandler


    End Sub


    Sub LockAnalRefTab(ByVal bool)

        frmH.panAnalRefStds.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp8.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 _
                     Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0 _
                )) Then

                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

        frmH.dgvCompanyAnalRef.Enabled = True
        frmH.dgvWatsonAnalRef.Enabled = True


        Dim gs As DataGridColumnStyle
        Dim Count1 As Short

        frmH.cmdAnalRefCancel.Enabled = Not (bool)
        frmH.cmdAddRepAnalyte.Enabled = Not (bool)
        frmH.cmdDeleteRepAnalyte.Enabled = Not (bool)
        frmH.cmdCopyRepAnalyte.Enabled = Not (bool)
        frmH.cmdAddAnalyte.Enabled = Not (bool)

        frmH.dgvCompanyAnalRef.ReadOnly = bool
        frmH.dgvCompanyAnalRef.Columns.Item(0).ReadOnly = True


    End Sub

    Sub LockCPTab(ByVal bool)

        frmH.panContr.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp9.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 _
                     Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0 _
                )) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

        frmH.dgvContributingPersonnel.Enabled = True

        frmH.cmdCPAdd.Enabled = Not (bool)
        frmH.cmdCPDelete.Enabled = Not (bool)
        frmH.cmdCPCancel.Enabled = Not (bool)
        frmH.gbMethodValMultiple.Enabled = Not (bool)
        frmH.cmdReplacePersonnel.Enabled = Not (bool)

        frmH.dgvContributingPersonnel.ReadOnly = bool

        If bool Then
        Else

        End If


    End Sub

    Sub LockHomeTab(ByVal bool)

        'bool=TRUE, lock the tab.   bool=FALSE, unlock the tab
        'LockHome has items that are not synchronous enabled
        'leave panChoose enabled                     
        frmH.panChoose.Enabled = True
        Dim ctrl As Control
        Dim var1
        Dim boolA As Boolean

        'first make all Action buttons disabled (except report history and show outstanding and Apply Template)
        For Each ctrl In frmH.panChoose.Controls

            Select Case UCase(ctrl.Name)
                Case UCase("panWatsonData")
                Case UCase("cmdReportHistory")
                Case UCase("cmdApplyTemplate")
                Case UCase("cmdShowOutstanding")
                Case UCase("cmdClearStudy")
                Case Else
                    Try
                        ctrl.Enabled = False
                    Catch ex As Exception
                    End Try
            End Select

            'If ((InStr(1, ctrl.Name, "panWatsonData", CompareMethod.Text) > 0) Or _
            '    (InStr(1, ctrl.Name, "cmdReportHistory", CompareMethod.Text) > 0) Or _
            '    (InStr(1, ctrl.Name, "cmdApplyTemplate", CompareMethod.Text) > 0) Or _
            '    (InStr(1, ctrl.Name, "cmdShowOutstanding", CompareMethod.Text) > 0)) Then
            'Else
            '    Try
            '        ctrl.Enabled = False
            '    Catch ex As Exception
            '    End Try
            'End If
        Next


        'then set all tab controls disabled (except labels)
        For Each ctrl In frmH.tp1.Controls

            If Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0) Then
                Try
                    ctrl.Enabled = False
                Catch ex As Exception
                End Try
            End If

        Next

        'now do things if allowed

        If BOOLHOME Then

            frmH.panChoose.Enabled = True
            For Each ctrl In frmH.panChoose.Controls
                If ((InStr(1, ctrl.Name, "panWatsonData", CompareMethod.Text) > 0) Or _
                (InStr(1, ctrl.Name, "cmdReportHistory", CompareMethod.Text) > 0) Or _
                (InStr(1, ctrl.Name, "cmdShowOutstanding", CompareMethod.Text) > 0)) Then  'Do nothing
                Else
                    Try
                        ctrl.Enabled = Not (bool)
                    Catch ex As Exception
                    End Try
                End If
            Next

            For Each ctrl In frmH.tp1.Controls

                If Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0) Then
                    Try
                        'Note: some are readonly setting
                        ctrl.Enabled = Not (bool)
                        Select Case ctrl.Name
                            Case "dgvReports"
                                frmH.dgvReports.ReadOnly = bool
                        End Select
                    Catch ex As Exception
                    End Try
                End If

            Next
            'these are available as bool  'NDL: Don't understand this.
            'cmdEdit/Cancel/Save calls need these control states as bool
            frmH.dgvwStudy.Enabled = bool
            frmH.gbStudyFilter.Enabled = bool
            frmH.cmdUpdateProject.Enabled = bool
            frmH.dgvReports.ReadOnly = bool

            If bool Then
                frmH.dgvwStudy.BackgroundColor = Color.White
                frmH.dgvwStudy.RowsDefaultCellStyle.BackColor = Color.White
            Else
                frmH.dgvwStudy.BackgroundColor = Color.Gray
                frmH.dgvwStudy.RowsDefaultCellStyle.BackColor = Color.Gray
            End If

            're-assess panStudyFilter
            frmH.optStudyDocStudies.Enabled = True
            frmH.optStudyDocOpen.Enabled = True
            frmH.optStudyDocClosed.Enabled = True

            frmH.cbxStudy.Enabled = bool
            frmH.MenuPrepareReport.Enabled = bool

            If frmH.cmdEdit.Enabled Then
                If frmH.rbOracle.Checked Then
                    frmH.panFilterStudy.Enabled = True
                    frmH.panStudyFilter.Enabled = True
                    frmH.gbStudyFilter.Enabled = True
                Else
                    frmH.panFilterStudy.Enabled = False
                    frmH.panStudyFilter.Enabled = False
                    frmH.gbStudyFilter.Enabled = False
                End If
            Else
                frmH.panFilterStudy.Enabled = False
                frmH.panStudyFilter.Enabled = False
                frmH.gbStudyFilter.Enabled = False
            End If

            'Hmmm. If 
            If frmH.rbOracle.Checked And boolFormLoad Then
                frmH.panFilterStudy.Enabled = True
                frmH.panStudyFilter.Enabled = True
                frmH.gbStudyFilter.Enabled = True

                'pesky
                Call SetPanAction()
            End If

            'the must always be shown
            frmH.cmdApplyTemplate.Enabled = True
            frmH.cmdClearStudy.Enabled = BOOLHOME

        End If



        'If bool Then
        'Else
        '    Dim dv As System.Data.DataView

        '    dv = frmH.dgvReports.DataSource
        '    If IsNothing(dv) Then
        '        frmH.cmdConfigureReport.Enabled = Not (bool)
        '    Else
        '        If dv.Count = 0 Then
        '            frmH.cmdConfigureReport.Enabled = Not (bool)
        '        Else
        '            frmH.cmdConfigureReport.Enabled = False
        '        End If
        '    End If


        '    frmH.cmdHomeCancel.Enabled = Not (bool)


        '    frmH.cmdCreateReportTitle.Enabled = Not (bool)
        '    frmH.cmdHomeCancel.Enabled = Not (bool)
        '    frmH.cmdApplyTemplate.Enabled = Not (bool)
        '    frmH.cmdHeader.Enabled = Not (bool)

        '    'frmh.cmdShowExample.Enabled = Not (bool)

        '    frmH.cmdLogin.Enabled = bool
        '    frmH.cmdChangePassword.Enabled = bool

        '    frmH.dgvReports.ReadOnly = bool

        '    frmH.dgvwStudy.Enabled = bool

        '    frmH.cbxStudy.Enabled = bool

        '    frmH.gbSource.Enabled = Not (bool)

        '    frmH.gbxMultVal.Enabled = Not (bool)
        'End If


    End Sub

    Sub LockReportStatementTab(ByVal bool As Boolean, boolDoA As Boolean)

        Dim var1


        'frmH.panWordTemp.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp5.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

        frmH.dgvReportStatements.Enabled = True
        frmH.dgvReportStatementWord.Enabled = True

        For Each ctrl In frmH.panWordTemp.Controls
            Try
                ctrl.Enabled = Not (bool)
            Catch ex As Exception

            End Try
        Next

        Dim a, b

        If (bool) Then
            frmH.lblWordStatements.Text = "Change to Edit mode to assign a different template."
            frmH.dgvReportStatementWord.ClearSelection()
            frmH.dgvReportStatementWord.Enabled = False
            frmH.dgvReportStatementWord.ForeColor = Color.Gray

        Else
            frmH.lblWordStatements.Text = "<< Doubleclick to assign Word Report Template"
            frmH.dgvReportStatementWord.Enabled = True
            frmH.dgvReportStatementWord.ForeColor = Color.Black
        End If

        a = frmH.lblWordStatements.Left
        b = frmH.lblWordStatements.Width

        frmH.gbxlblChooseEditWordTemplate.Width = a + b + 10

        frmH.dgvReportStatements.ReadOnly = bool
        frmH.cmdCancelReportStatements.Enabled = Not (bool)
        frmH.cmdOrderReportBodySection.Enabled = Not (bool)
        frmH.cmdRBSAll.Enabled = Not (bool)

        Try
            frmH.dgvReportStatements.Columns.Item("charStatement").ReadOnly = True
            frmH.dgvReportStatements.Columns.Item("charSectionName").ReadOnly = True
        Catch ex As Exception

        End Try



        'NOTE:  Edit and View Templates have their own permissions
        'make them enabled no matter what
        '20170622 LEE: No, should be disabled if in edit mode

        frmH.cmdOpenReportStatements.Enabled = bool ' True 'BOOLEDITWORDTEMPLATE
        frmH.cmdRefreshStatements.Enabled = bool ' True 'BOOLVIEWWORDTEMPLATE
        '



    End Sub

    Sub LockAdministration(ByVal bool)

        Dim ctrl As Control
        For Each ctrl In frmH.tp14.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then
            Else
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If

        Next

    End Sub



    Sub LockSampleReceiptTab(ByVal bool)

        frmH.panSampleRec.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp12.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 _
                     Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0 _
                )) Then

                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

        frmH.dgvSampleReceipt.Enabled = True
        frmH.dgvSampleReceiptWatson.Enabled = True

        frmH.dgvSampleReceipt.ReadOnly = bool
        frmH.cmdInsertSRec.Enabled = Not (bool)
        frmH.cmdDeletSRec.Enabled = Not (bool)
        frmH.cmdSRecCancel.Enabled = Not (bool)

        frmH.chkUseWatsonSampleNumber.Enabled = Not (bool)
        frmH.chkManualSampleNumber.Enabled = Not (bool)

        If bool Then
            frmH.txtSRecTotalReport.ReadOnly = True
        Else
            If frmH.chkManualSampleNumber.CheckState = CheckState.Checked Then
                frmH.txtSRecTotalReport.ReadOnly = False
            End If
        End If

        If bool Then

        Else

        End If


    End Sub

    Sub LockReportGeneration(ByVal bool As Boolean)

        frmH.cbxExampleReport.Enabled = Not (bool)

    End Sub

    Sub LockDataTab(ByVal bool)

        Dim str2 As String
        Dim dgv As DataGridView
        Dim Count1 As Short

        frmH.panTopLevel.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp2.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then
                Try  ' For everything except labels
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception
                End Try
            End If
        Next

        frmH.tabData.Enabled = True

        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim int1 As Short

        frmH.cbxAnticoagulant.Enabled = Not (bool)
        frmH.cbxAssayTechnique.Enabled = Not (bool)
        frmH.cbxAssayTechniqueAcronym.Enabled = False ' Not (bool)
        frmH.cbxSampleSizeUnits.Enabled = Not (bool)
        frmH.cbxSampleStorageTemp.Enabled = Not (bool)
        frmH.cbxSubmittedTo.Enabled = Not (bool)
        frmH.cbxSubmittedBy.Enabled = Not (bool)
        frmH.cbxInSupportOf.Enabled = Not (bool)

        frmH.gbRound5.Enabled = Not (bool)
        frmH.gbCritPrecision.Enabled = Not (bool)
        frmH.gbMeanComp.Enabled = Not (bool)

        frmH.cmdDataCancel.Enabled = Not (bool)

        frmH.dgvDataCompany.ReadOnly = bool

        int1 = frmH.dgvDataCompany.Columns.Count 'debug

        frmH.dgvDataCompany.Columns.Item(0).ReadOnly = True
        frmH.dgvDataCompany.Columns.Item("Example").ReadOnly = True

        Try
            frmH.dgvStudyConfig.ReadOnly = bool
            frmH.dgvStudyConfig.Columns.Item(0).ReadOnly = True
            frmH.dgvStudyConfig.Columns.Item("Example").ReadOnly = True
        Catch ex As Exception

        End Try

        Try
           
            dgv = frmH.dgvFC

            dv = frmH.dgvFC.DataSource
            If bool Then
                dv.AllowDelete = False
                dv.AllowEdit = False
                dv.AllowNew = False
                dgv.ReadOnly = True
            Else
                dv.AllowEdit = True
                dgv.ReadOnly = False
            End If

            str2 = "CHARFIELDCODE"
            dgv.Columns(str2).ReadOnly = True
            str2 = "CHARDESCRIPTION"
            dgv.Columns(str2).ReadOnly = True
            str2 = "CHAREXAMPLE"
            dgv.Columns(str2).ReadOnly = True
            str2 = "CHARVALUE"
            dgv.Columns(str2).ReadOnly = bool

        Catch ex As Exception

        End Try

        '20180828 LEE:
        'make some columns editable
        Try

            dgv = frmH.dgvAnalyteGroups
            dgv.ReadOnly = bool

            For count1 = 0 To frmH.dgvAnalyteGroups.Columns.Count - 1
                dgv.Columns(Count1).ReadOnly = True
            Next

            str2 = "CHARUSERANALYTE"
            dgv.Columns(str2).ReadOnly = bool
            str2 = "CHARUSERIS"
            dgv.Columns(str2).ReadOnly = bool


        Catch ex As Exception
            str2=str2
        End Try

        If bool Then
        Else

        End If



    End Sub

    Sub LockReportTableHeaderConfigTab(ByVal bool)

        frmH.panColHeadings.Enabled = Not (bool)
        Dim ctrl As Control
        For Each ctrl In frmH.tp7.Controls
            If (Not (InStr(1, ctrl.Name, "lbl", CompareMethod.Text) > 0 Or InStr(1, ctrl.Name, "Label", CompareMethod.Text) > 0)) Then
                Try
                    ctrl.Enabled = Not (bool)
                Catch ex As Exception

                End Try
            End If
        Next

        frmH.dgvReportTableConfiguration.Enabled = True
        frmH.dgvReportTables.Enabled = True
        frmH.dgvReportTableHeaderConfig.Enabled = True 'So we can see the ToolTips even when it's locked (it is still read-only)



        'frmh.cbxReportTableTypesa.Enabled = Not (bool)

        frmH.cmdRTHeaderConfigCancel.Enabled = Not (bool)

        frmH.dgvReportTableHeaderConfig.ReadOnly = bool

        If bool Then

        Else

        End If



    End Sub

    Sub DoRTConfigCancel()

        Call FillReportTableHome()
        Call ReportsSelection(True)
        Call FillTableReports(False)
        Call FillTableReportsAnalytes(True)
        Call FillTableReportDataAnalytes(False)
        'pesky
        Call OrderReportTableConfig()

        Call AssessSampleAssignment()

    End Sub

    'Sub ConfigReportStyle()

    '    Dim dv As System.Data.DataView
    '    Dim int1 As Short
    '    Dim int2 As Short

    '    Call FillTableReports(False)
    '    Call FillTableReportsAnalytes(True)
    '    Call FillTableReportDataAnalytes(False)
    '    'pesky
    '    Call OrderReportTableConfig()

    '    'boolCont = False '???
    '    boolCont = True

    'End Sub

    Sub ReportsSelection(ByVal boolCancel)

        Dim int1 As Short
        Dim int2 As Short
        Dim dv As System.Data.DataView
        Dim str1 As String
        Dim strF As String
        Dim var1, var2, var3, var4
        Dim Count1 As Short
        Dim drows() As DataRow

        If frmH.dgvReports.CurrentRow Is Nothing Then
            id_tblReports = 0
            Exit Sub
        End If
        int1 = frmH.dgvReports.CurrentRow.Index
        If int1 = -1 Then 'no reports configured or selected
            id_tblReports = 0
            Exit Sub
        End If
        If boolCancel Then
            dv = frmH.dgvReports.DataSource
            'get id_tblConfigReportType
            var4 = dv.Item(int1).Item("id_tblReports")
            id_tblReports = var4
            'get values for Report Config tab group boxes from underlying table and enter the values
            Dim tbl As System.Data.DataTable
            tbl = tblReports
            strF = "id_tblReports = " & var4
            drows = tbl.Select(strF)
            'var1 = drows(0).Item("id_tblReports")
            var1 = drows(0).Item("id_tblConfigReportType")
        Else
            'get values for Report Config tab group boxes from dv and enter the values
            dv = frmH.dgvReports.DataSource
            'get id_tblConfigReportType
            var1 = dv.Item(int1).Item("id_tblConfigReportType")
        End If

        If IsDBNull(var1) Then 'set default as sample report
            var1 = 1
        End If


        'get intUserSummary
        If boolCancel Then
            var3 = drows(0).Item("intUserComments")
        Else
            var3 = dv.Item(int1).Item("intUserComments")
        End If
        var3 = NZ(var3, 1)
        Select Case var3
            Case 1
                frmH.rbUseWatsonComments.Checked = True
            Case 2
                frmH.rbUseUserComments.Checked = True
        End Select

        ''get boolexcludepsae
        ''20160907 LEE: boolexcludepsae deprecated
        'If boolCancel Then
        '    var3 = drows(0).Item("boolexcludepsae")
        'Else
        '    var3 = dv.Item(int1).Item("boolexcludepsae")
        'End If
        'var3 = NZ(var3, -1)

        'Dim bool As Boolean
        'bool = boolStopRBS
        'boolStopRBS = True
        'Select Case var3
        '    Case -1
        '        frmH.rbAnalRunsExclPSAE.Checked = True
        '        frmH.rbAnalRunsShowAll.Checked = False
        '    Case 0
        '        frmH.rbAnalRunsExclPSAE.Checked = False
        '        frmH.rbAnalRunsShowAll.Checked = True
        'End Select
        'boolStopRBS = bool

        'BOOLALLAR chkAll
        'BOOLACCAR chkAccepted
        'BOOLREJAR chkRejected
        'BOOLREGRAR chkRegrPerformed
        'BOOLNOREGRAR chkNoRegrPerformed
        'BOOLINCLPSAE chkPSAE

        Dim bool As Boolean
        bool = boolStopRBS
        boolStopRBS = True

        If boolCancel Then
            var3 = drows(0).Item("BOOLALLAR")
        Else
            var3 = dv.Item(int1).Item("BOOLALLAR")
        End If
        var3 = NZ(var3, -1)
        frmH.chkAll.Checked = var3 'checkbox will take 1 or -1 for checked

        If boolCancel Then
            var3 = drows(0).Item("BOOLACCAR")
        Else
            var3 = dv.Item(int1).Item("BOOLACCAR")
        End If
        var3 = NZ(var3, 0)
        frmH.chkAccepted.Checked = var3

        If boolCancel Then
            var3 = drows(0).Item("BOOLREJAR")
        Else
            var3 = dv.Item(int1).Item("BOOLREJAR")
        End If
        var3 = NZ(var3, 0)
        frmH.chkRejected.Checked = var3

        If boolCancel Then
            var3 = drows(0).Item("BOOLREGRAR")
        Else
            var3 = dv.Item(int1).Item("BOOLREGRAR")
        End If
        var3 = NZ(var3, 0)
        frmH.chkRegrPerformed.Checked = var3

        If boolCancel Then
            var3 = drows(0).Item("BOOLNOREGRAR")
        Else
            var3 = dv.Item(int1).Item("BOOLNOREGRAR")
        End If
        var3 = NZ(var3, 0)
        frmH.chkNoRegrPerformed.Checked = var3

        If boolCancel Then
            var3 = drows(0).Item("BOOLINCLPSAE")
        Else
            var3 = dv.Item(int1).Item("BOOLINCLPSAE")
        End If
        var3 = NZ(var3, 0)
        frmH.chkPSAE.Checked = var3

        boolStopRBS = bool

    End Sub

    Sub CARcellChanged()

        Exit Sub


        Dim var1, var2
        'Dim dg As DataGrid
        Dim dgv As DataGridView
        Dim intRowDate1 As Short
        Dim intRowDate2 As Short
        Dim intRowBool As Short
        Dim tbl As System.Data.DataTable
        Dim intRow As Short
        Dim intCol As Short
        Dim varNull As System.DBNull
        Dim str1 As String
        Dim dt As Date

        tbl = tblCompanyAnalRefTable
        'dg = frmh.dgCompanyAnalRef
        dgv = frmH.dgvCompanyAnalRef
        intRowDate1 = FindRow("Date Received", tbl, "Item")
        intRowDate2 = FindRow("Expiration/Retest Date", tbl, "Item")
        intRowBool = FindRow("Is Replicate?", tbl, "Item")
        If oldCurrentRowCAR = -1 Or oldCurrentColCAR = -1 Then
            Exit Sub
        End If

        newCurrentCellCAR = NZ(dgv.Item(oldCurrentRowCAR, oldCurrentColCAR), varNull)
        If (intRowDate1 = oldCurrentRowCAR Or intRowDate2 = oldCurrentRowCAR) And oldCurrentColCAR <> 0 Then
            If Len(NZ(newCurrentCellCAR, "")) = 0 Then
                boolOKtoVal = True
            ElseIf IsDate(newCurrentCellCAR) Then
                boolOKtoVal = True
                If StrComp(newCurrentCellCAR, oldCurrentCellCAR, CompareMethod.Text) = 0 Then
                Else
                    dt = newCurrentCellCAR
                    'format data to a consistent standard
                    str1 = Format(dt, LDateFormat)
                    str1 = CStr(dt)
                    dgv.Item(oldCurrentRowCAR, oldCurrentColCAR).Value = str1
                End If
            Else
                valErr = "This value must be in an acceptable date format."
                boolOKtoVal = False
            End If
        ElseIf intRowBool = oldCurrentRowCAR Then
            'Value must be Yes or No
            'If StrComp(newCurrentCellCAR, "Yes", CompareMethod.Text) = 0 Or StrComp(newCurrentCellCAR, "No", CompareMethod.Text) = 0 Then
            If StrComp(newCurrentCellCAR, oldCurrentCellCAR, CompareMethod.Text) = 0 Then ' Or StrComp(newCurrentCellCAR, "No", CompareMethod.Text) = 0 Then
                boolOKtoVal = True
            Else
                valErr = "This value cannot be modified."
                boolOKtoVal = False
            End If
        End If
        If boolOKtoVal Then
            oldCurrentRowCAR = dgv.CurrentCell.RowIndex
            oldCurrentColCAR = dgv.CurrentCell.ColumnIndex
            oldCurrentCellCAR = NZ(dgv.Item(oldCurrentRowCAR, oldCurrentColCAR).Value, "")
            'MsgBox("R" & oldCurrentRowCAR & ":C" & oldCurrentColCAR & " = " & oldCurrentCellCAR)
        Else
            boolFromCAR = True
            dgv.CurrentCell = dgv(oldCurrentRowCAR, oldCurrentColCAR)

            boolFromCAR = False
            MsgBox(valErr, MsgBoxStyle.Information, "Validation Error...")
            dgv.Item(oldCurrentRowCAR, oldCurrentColCAR).Value = oldCurrentCellCAR

        End If
    End Sub

    Sub doMethValCancel()

        'reject changes to tables
        tblMethodValidationData.RejectChanges()
        'tblMethodValData.RejectChanges()
        tblMethValExistingGuWu.RejectChanges()


        Call FillMethValExistingGuWu()

    End Sub

    Sub MethValMultipleAddColumns(ByVal intAdd As Short)

        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim intRows As Short
        Dim intCols As Short
        Dim dv As System.Data.DataView
        Dim str1 As String
        Dim var1, var2

        If boolLoad Then
            Exit Sub
        End If
        dtbl = tblMethodValData
        dgv = frmH.dgvMethodValData
        'ts1 = dg.TableStyles(0)
        'gc = ts1.GridColumnStyles
        intRows = dtbl.Rows.Count
        intCols = dtbl.Columns.Count

        For Count1 = 0 To intCols - 1 'debugging
            var1 = dtbl.Columns(Count1).ColumnName
            var2 = var1
        Next

        'delete all but two tbl columns 
        For Count1 = intCols - 1 To 2 Step -1
            dtbl.Columns.Remove(dtbl.Columns.Item(Count1))
        Next
        'delete any data from column 1
        For Count1 = 0 To intRows - 1
            dtbl.Rows.Item(Count1).Item(1) = ""
        Next

        Select Case intAdd
            Case 0 'revert to a single method
                'ts1.GridColumnStyles(1).HeaderText = arrAnalytes(1, 1) '"Value"
                dgv.Columns(1).HeaderText = arrAnalytes(1, 1)
                var1 = arrAnalytes(1, 1)
                dtbl.Columns.Item(1).ColumnName = arrAnalytes(1, 1)
                dtbl.Columns.Item(1).ReadOnly = False

                ''update gridstyle
                'Dim gc1 As DataGridTextBoxColumn
                'gc1 = ts1.GridColumnStyles(1)
                'gc1.MappingName = arrAnalytes(1, 1)
                'gc1.HeaderText = arrAnalytes(1, 1)
                'gc1.NullText = ""
                'gc1.ReadOnly = False

            Case 1 'configure for number of analytes
                If ctAnalytes = 0 And boolFormLoad = False Then
                    MsgBox("This StudyDoc study doesn't have any configured analytes.", MsgBoxStyle.Information, "No analytes...")
                    'ts1.GridColumnStyles(1).HeaderText = "Value"
                    dgv.Columns(1).HeaderText = "Value"
                    dtbl.Columns.Item(1).ColumnName = "Value"
                    dtbl.Columns.Item(1).ReadOnly = False

                    'frmH.chkMethodValMultiple.Checked = False
                    'frmH.txtMethValMultiple.Enabled = False
                    'frmH.cmdMethValExecute.Enabled = False

                    ''update gridstyle
                    'Dim gc1 As DataGridTextBoxColumn
                    'gc1 = ts1.GridColumnStyles(1)
                    'gc1.MappingName = "Value"
                    'gc1.HeaderText = "Value"
                    'gc1.NullText = ""
                    'gc1.ReadOnly = False


                Else
                    For Count1 = 1 To ctAnalytes
                        'add column
                        If Count1 = 1 Then
                            Dim col As DataColumn
                            col = dtbl.Columns.Item(1)
                            col.ColumnName = arrAnalytes(1, Count1)
                            col.Caption = arrAnalytes(1, Count1)
                            col.DataType = System.Type.GetType("System.String")
                            col.ReadOnly = False
                            col.AllowDBNull = True
                            str1 = col.ColumnName
                            dgv.Columns(Count1).HeaderText = str1
                        Else
                            Dim col As New DataColumn
                            col.ColumnName = arrAnalytes(1, Count1)
                            col.Caption = arrAnalytes(1, Count1)
                            col.DataType = System.Type.GetType("System.String")
                            col.ReadOnly = False
                            col.AllowDBNull = True
                            dtbl.Columns.Add(col)
                            str1 = col.ColumnName

                            dgv.Columns(Count1).HeaderText = str1
                            'add gridstyle
                            'If Count1 = 1 Then
                            '    Dim gc1 As DataGridTextBoxColumn
                            '    gc1 = ts1.GridColumnStyles(1)
                            '    gc1.MappingName = str1
                            '    gc1.HeaderText = str1
                            '    gc1.NullText = ""
                            '    gc1.ReadOnly = False
                            'Else
                            '    Dim gc1 As New DataGridTextBoxColumn
                            '    gc1.MappingName = str1
                            '    gc1.HeaderText = str1
                            '    gc1.NullText = ""
                            '    gc1.Width = ts1.GridColumnStyles(1).Width
                            '    gc1.ReadOnly = False
                            '    ts1.GridColumnStyles.Add(gc1)
                            'End If
                        End If
                    Next
                End If
            Case 2 'this won't happen anymore
                'var1 = frmh.txtMethValMultiple.Text
                'var2 = CInt(var1)
                'For Count1 = 1 To var2
                '    'add column
                '    If Count1 = 1 Then
                '        Dim col As DataColumn
                '        col = dtbl.Columns.item(1)
                '        col.ColumnName = arrAnalytes(1, 1) & "_" & Count1 '"Value" & Count1
                '        col.Caption = arrAnalytes(1, 1) & "_" & Count1 ' "Value" & Count1
                '        col.DataType = System.Type.GetType("System.String")
                '        str1 = col.ColumnName
                '    Else
                '        Dim col As New DataColumn
                '        col.ColumnName = arrAnalytes(1, 1) & "_" & Count1 '"Value" & Count1
                '        col.Caption = arrAnalytes(1, 1) & "_" & Count1 '"Value" & Count1
                '        col.DataType = System.Type.GetType("System.String")
                '        dtbl.Columns.Add(col)
                '        str1 = col.ColumnName
                '    End If

                '    'add gridstyle
                '    If Count1 = 1 Then
                '        Dim gc1 As DataGridTextBoxColumn
                '        gc1 = ts1.GridColumnStyles(1)
                '        gc1.MappingName = str1
                '        gc1.HeaderText = str1
                '        gc1.NullText = ""
                '        gc1.ReadOnly = False
                '    Else
                '        Dim gc1 As New DataGridTextBoxColumn
                '        gc1.MappingName = str1
                '        gc1.HeaderText = str1
                '        gc1.NullText = ""
                '        gc1.Width = ts1.GridColumnStyles(1).Width
                '        gc1.ReadOnly = False
                '        ts1.GridColumnStyles.Add(gc1)
                '    End If
                'Next
        End Select
        dv = dtbl.DefaultView
        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = True
        dgv.DataSource = dv

        Call MethValAutoCol()


end1:

    End Sub

    Sub FillMethValIfVal()

        '20190208 LEE:
        'this get's called when Sample Analysis has a linked Method Validation study
        'as apposed to FillTableStuffMethVal

        'this routine will fill dgvMethodValData
        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim dvS As System.Data.DataView
        Dim dgvR As DataGridView


        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim Count1 As Short
        Dim int1 As Short
        Dim var1, var2
        Dim id As Int64
        Dim boolMethVal As Boolean
        Dim intCols As Short
        Dim intRows As Short
        Dim idMax As Int64
        Dim boolM As Boolean

        dgvS = frmH.dgvDataCompany
        dgvD = frmH.dgvMethodValData
        dgvR = frmH.dgvReports
        dvS = dgvS.DataSource

        intCols = ctAnalytes

        boolMethVal = False
        If dgvR.Rows.Count = 0 Then 'skip
        Else
            'find if meth validation
            id = NZ(dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value, 1)
            If id > 1 And id < 5 Then
                boolMethVal = True
            Else
                boolMethVal = False
            End If
        End If

        If boolMethVal Then
        Else
            GoTo end1
        End If

        dtbl = tblMethodValidationData
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        rows = dtbl.Select(strF)
        intRows = rows.Length

        Dim str1 As String
        Dim str2 As String
        Dim Count2 As Short
        Dim boolHit As Boolean
        Dim intH As Short

        'If intRows <> ctAnalytes Then 'need to add rows
        intH = 0
        If intRows < ctAnalytes Then 'need to add rows
            'idMax = GetMaxID("tblMethodValidationData")
            For Count1 = 1 To ctAnalytes

                'look for cmdpd
                str1 = tblAnalytesHome.Rows(Count1 - 1).Item("AnalyteDescription")
                boolHit = False
                For Count2 = 0 To rows.Length - 1
                    str2 = rows(Count2).Item("CHARCOLUMNNAME")
                    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                        boolHit = True
                        Exit For
                    End If
                Next

                If boolHit Then
                    intH = intH + 1
                Else
                    intH = intH + 1
                    idMax = idMax + 1
                    Dim nrow As DataRow = dtbl.NewRow
                    nrow.BeginEdit()
                    nrow("ID_TBLSTUDIES") = id_tblStudies ' idMax
                    nrow("ID_TBLSTUDIES2") = 0
                    nrow("INTCOLUMNNUMBER") = intH ' Count1
                    nrow("ID_TBLREPORTS") = 0 'default for now
                    nrow("CHARCOLUMNNAME") = str1
                    nrow.EndEdit()
                    'dtbl.Rows.Add(nrow)
                    Try
                        dtbl.Rows.Add(nrow)
                    Catch ex As Exception
                        var1 = "A"
                    End Try
                End If
            Next
            'boolM = PutMaxID("tblMethodValidationData", idMax)

            strF = "ID_TBLSTUDIES = " & id_tblStudies
            Erase rows
            rows = dtbl.Select(strF)
            intRows = rows.Length

        End If

        For Count1 = 0 To rows.Length - 1

            rows(Count1).BeginEdit()
            'int1 = FindRowDV("Corporate Study/Project Number", dvS)
            'var1 = dgvS(1, int1).Value
            int1 = FindRowDV("Corporate Study/Project Number", dvS)
            var1 = dgvS(1, int1).Value
            rows(Count1).Item("CHARCORPORATESTUDYID") = var1

            int1 = FindRowDV("Protocol Number", dvS)
            var1 = dgvS(1, int1).Value
            rows(Count1).Item("CHARPROTOCOLNUMBER") = var1

            int1 = FindRowDV("Sponsor Study Number", dvS)
            var1 = dgvS(1, int1).Value
            rows(Count1).Item("CHARSPONSORMETHODVALIDATIONID") = var1

            int1 = FindRowDV("Sponsor Study Title", dvS)
            var1 = dgvS(1, int1).Value
            rows(Count1).Item("CHARSPONSORMETHVALTITLE") = var1

            If dgvR.Rows.Count = 0 Then
            Else

                rows(Count1).Item("CHARMETHODVALIDATIONTITLE") = dgvR("CHARREPORTTITLE", 0).Value

                rows(Count1).Item("CHARVALREPORTNUM") = dgvR("CHARREPORTNUMBER", 0).Value

            End If

            rows(Count1).Item("CHARANALMETHODTYPE") = frmH.cbxAssayTechniqueAcronym.SelectedItem


            rows(Count1).EndEdit()
        Next

end1:

    End Sub

    Sub FillMethValExistingGuWu()

        'called by
        '  UpdateProjectClick063
        '  GetStudyInfo – repeated at end of action because was not actuating during study load
        '  RealMethValExecute
        '  doMethValCancel




        Dim tbl As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim tbl_1 As System.Data.DataTable
        Dim tb As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim ts1 As DataGridTableStyle
        Dim col As DataColumn
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String
        Dim str1 As String
        Dim str2 As String
        Dim drows() As DataRow
        Dim drow As DataRow
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim intRows As Short
        Dim intRows1 As Short
        Dim var1, var2, var3, var4
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intAdd As Short
        Dim tblRowD As System.Data.DataTable
        Dim drowRowD() As DataRow
        Dim intRowD As Short
        Dim int1 As Short
        Dim dr() As DataRow
        Dim dvM As System.Data.DataView
        Dim var8
        Dim strAnal As String

        intAdd = 1
        Call MethValMultipleAddColumns(intAdd)

        'if method validation, then overwrite from data
        Call FillMethValIfVal()

        tbl_1 = tblMethValExistingGuWu
        dtbl = tblMethodValidationData
        strF = "id_tblStudies = " & id_tblStudies
        strS = "" ' "CHARCOLUMNNAME ASC"
        drows = dtbl.Select(strF, strS)
        intRows = drows.Length
        dgv1 = frmH.dgvMethodValData
        dgv2 = frmH.dgvMethValExistingGuWu


        tbl = tblMethodValData
        intRows1 = tbl.Columns.Count - 1

        'add data to tbl
        If intRows = 0 Then 'no data saved
        Else
            strF = "charDataTableName = 'tblMethodValidationData'"
            tblRowD = tblDataTableRowTitles
            drowRowD = tblRowD.Select(strF, "intOrder ASC")
            intRowD = drowRowD.Length
            For Count1 = 0 To intRows - 1 'this is # of cmpd sets
                'If Count1 > intRows1 - 1 Then
                '    Exit For
                'End If

                strAnal = drows(Count1).Item("CHARCOLUMNNAME")
                '20171108 LEE: get from tblAnalyteGroups
                strF1 = "ANALYTEDESCRIPTION_C = '" & CleanText(strAnal) & "'"
                '20190206 LEE:
                'Rearrange for debugging. Sometimes analyte has a prime (') in it (Alturas ONT380005)
                'Datatable select statement doesn't like this
                'need to adapt to (') in analyte name
                Dim rowsAG() As DataRow
                rowsAG = tblAnalyteGroups.Select(strF1)


                For Count2 = 0 To intRowD - 1
                    str1 = NZ(drowRowD(Count2).Item("charTableRefColumnName"), "")
                    tbl.Rows.Item(Count2).BeginEdit()
                    'drowRowD(Count2).BeginEdit()
                    If Len(str1) = 0 Then
                        tbl.Rows.Item(Count2).Item(Count1 + 1) = ""
                        'drowRowD(Count2).Item(Count1 + 1) = ""
                    Else
                        var1 = NZ(drows(Count1).Item(str1), "")
                        Select Case str1
                            Case "CHARSPECIES"
                                If Len(var1) = 0 Or StrComp(var1, "[NA]", CompareMethod.Text) = 0 Then 'get from Watson data
                                    dvM = frmH.dgvDataWatson.DataSource 'intI is analyte column in dgvMethodValData
                                    int1 = FindRowDVByCol("Species", dvM, "Item")
                                    var1 = Trim(NZ(dvM.Item(int1).Item(1), "[NA]"))
                                End If
                            Case "CHARMATRIX"
                                '20171108 LEE: get from tblAnalyteGroups
                                If rowsAG.Length = 0 Then
                                    var1 = "NA"
                                Else
                                    var1 = NZ(rowsAG(0).Item("MATRIX"), "NA")
                                End If
                                var1 = var1 'debug

                                'If Len(var1) = 0 Or StrComp(var1, "[NA]", CompareMethod.Text) = 0 Then 'get from Watson data
                                '    dvM = frmH.dgvDataWatson.DataSource 'intI is analyte column in dgvMethodValData
                                '    int1 = FindRowDVByCol("Matrix", dvM, "Item")
                                '    var1 = Trim(NZ(dvM.Item(int1).Item(1), "[NA]"))
                                'End If

                            Case "NUMSAMPLESIZE"

                                '20171108 LEE: get from tblSpeciesMatrix
                                '20180330 LEE: Do this only if existing value is null or blank
                                If Len(var1) = 0 Then

                                    'find matrix
                                    If rowsAG.Length = 0 Then
                                        var2 = "NA"
                                    Else
                                        var2 = NZ(rowsAG(0).Item("MATRIX"), "NA")
                                    End If
                                    strF2 = "SAMPLETYPEID = '" & var2 & "'"
                                    '2017111 LEE:
                                    'Dim rowsSM() As DataRow = tblSpeciesMatrix.Select(strF2)
                                    Dim rowsSM() As DataRow = tblSpeciesMatrixSV.Select(strF2)
                                    If rowsSM.Length = 0 Then
                                        var1 = 0
                                    Else
                                        For Count3 = 0 To rowsSM.Length - 1
                                            var3 = NZ(rowsSM(Count3).Item("STANDARDVOLUME"), "")
                                            If IsNumeric(var3) Then
                                                var1 = var3
                                                Exit For
                                            Else
                                                var1 = 0
                                            End If
                                        Next
                                    End If

                                End If

                            Case Else

                        End Select
                        'here's a problem
                        'tbl.Rows.Item(Count2).Item(Count1 + 1) = var1 ' NZ(drows(Count1).Item(str1), "")
                        Try
                            '20171107 LEE:
                            tbl.Rows.Item(Count2).Item(Count1 + 1) = NZ(var1, DBNull.Value) 'var1  ' NZ(drows(Count1).Item(str1), "")
                            'drowRowD(Count2).Item(Count1 + 1) = var1
                        Catch ex As Exception

                        End Try
                    End If
                    'drowRowD(Count2).EndEdit()
                    tbl.Rows.Item(Count2).EndEdit()
                Next
            Next
        End If

        'assign dv.tbl to dgv1
        'dv = tbl.DefaultView
        dv = New DataView(tbl)
        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = True
        dgv1.DataSource = dv

        'format dg2
        'check to see contents of tbl
        'tbl_1.Rows.Clear()
        tbl_1.Clear()

        Dim boolV As Boolean = False

        If intRows = 0 Then '  'no data saved yet,so add default stuff
            For Count1 = 1 To intRows1
                drow = tbl_1.NewRow
                drow.BeginEdit()
                var1 = tbl.Columns.Item(Count1).ColumnName
                'var1 = frmH.dgvMethodValData.TableStyles(0).GridColumnStyles(Count1).HeaderText
                var1 = frmH.dgvMethodValData.Columns(Count1).HeaderText
                drow("ColumnName") = var1
                'leave last three columns blank
                drow.EndEdit()
                tbl_1.Rows.Add(drow)
            Next
            frmH.gbxlblReviewValidatedMethod.Visible = False

        Else 'enter info from dtbl

            tb = tblStudies
            For Count1 = 1 To intRows ' intRows1
                int1 = NZ(drows(Count1 - 1).Item("id_tblStudies2"), 0)
                If int1 = 0 Then 'add a blank row

                    drow = tbl_1.NewRow
                    drow.BeginEdit()
                    drow("ColumnName") = drows(Count1 - 1).Item("charColumnName")

                    var3 = NZ(drows(Count1 - 1).Item("id_tblStudies2"), "")
                    If Len(CStr(var3)) = 0 Then
                    Else
                        If var3 = 0 Then
                        Else
                            boolV = True
                        End If

                    End If

                    drow("id_tblStudies") = drows(Count1 - 1).Item("id_tblStudies2")
                    drow("CHARARCHIVEPATH") = drows(Count1 - 1).Item("CHARARCHIVEPATH")
                    'find Watson Study
                    'strF = "id_tblStudies = " & drows(Count1 - 1).Item("id_tblStudies2")
                    'dr = tb.Select(strF)
                    'drow("Watson Study") = dr(0).Item("charWatsonStudyName")
                    drow.EndEdit()
                    tbl_1.Rows.Add(drow)

                Else
                    drow = tbl_1.NewRow
                    drow.BeginEdit()
                    drow("ColumnName") = drows(Count1 - 1).Item("charColumnName")
                    drow("id_tblStudies") = drows(Count1 - 1).Item("id_tblStudies2")
                    'find Watson Study
                    var1 = NZ(drows(Count1 - 1).Item("id_tblStudies2"), "")
                    If Len(var1) = 0 Or var1 = 0 Then
                    Else
                        boolV = True
                    End If
                    strF = "id_tblStudies = " & drows(Count1 - 1).Item("id_tblStudies2")
                    dr = tb.Select(strF)
                    drow("WatsonStudy") = dr(0).Item("charWatsonStudyName")
                    drow("CHARARCHIVEPATH") = drows(Count1 - 1).Item("CHARARCHIVEPATH")
                    drow.EndEdit()
                    tbl_1.Rows.Add(drow)
                End If
            Next
        End If
        'dv = tbl_1.DefaultView
        Dim dv1 As System.Data.DataView = New DataView(tblMethValExistingGuWu)
        dv1.AllowNew = False
        dv1.AllowDelete = False
        dv1.AllowEdit = True
        'strS = "ColumnName ASC"
        'dv1.Sort = strS
        dgv2.DataSource = dv1

        dgv2.AutoResizeColumns()
        dgv2.AutoResizeRows()

        'frmH.gbxlblReviewValidatedMethod.Visible = boolV
        '20190220 LEE:
        frmH.gbxlblReviewValidatedMethod.Visible = True


        'get more stuff
        Dim dgvR As DataGridView
        Dim idR As Int64
        dgvR = frmH.dgvReports
        If dgvR.Rows.Count = 0 Then
        Else
            idR = NZ(dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value, 0)
            If idR > 1 And idR < 1000 Then
                Call FillTableStuffMethVal(False)
            End If
        End If

        'must save all this stuff' NO!! screws up Audit Trail
        'Call SaveMethValTab()

    End Sub

End Module
