Option Compare Text

Public Class frmReportTableConfig

    Public intORow As Short
    Public boolHold As Boolean = False
    Public boolFormLoad As Boolean = True
    Public boolFromEdit As Boolean = False
    Public arrBU(0, 0)
    Public charOrigLegend As String = ""
    Public idSel As Int64
    Public boolTS As Boolean = False
    Public gidTR As Int64
    Public gidCRT As Int64
    Public boolConfigISR As Boolean = False

    Public strFilter As String = ""

    Private tblSAS As New System.Data.DataTable ' DataTable

    Public boolClose As Boolean = False

    Sub Check_ID_TblReportTable()

        '20181218 LEE:
        'Don't know how this happened, but TBLTABLEPROPERTIES isn't in sync with TBLREPORTTABLE in Frontage study BTM-2421
        'a cursory check of other studies doesn't show this
        'will run this code at the opening of Report Table Advanced Configuration just in case
        'Pretty sure this occurred when an Apply Template was performed
        'will not use for now

        Exit Sub

        Dim dtblRT As DataTable = tblReportTable
        Dim dtblTP As DataTable = tblTableProperties

        Dim strF As String
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim strF1 As String
        Dim strF2 As String

        Dim rowsRT() As DataRow = dtblRT.Select(strF)
        Dim rowsTP() As DataRow = dtblTP.Select(strF)
        Dim intRowsRT As Int16 = rowsRT.Length
        Dim intRowsTP As Int16 = rowsTP.Length

        Dim idRT As Int32
        Dim intMax As Int32
        Dim idCRT As Int32


        If intRowsRT = intRowsTP Then
            GoTo end1
        End If

        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim var1, var2, var3, var4

        Dim boolHit As Boolean = False

        For Count1 = 0 To intRowsRT - 1

            idRT = rowsRT(Count1).Item("ID_TBLREPORTTABLE")
            strF1 = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idRT
            Dim rows1() As DataRow = dtblTP.Select(strF1)

            If rows1.Length > 0 Then
            Else

                If boolHit = False Then
                    'get maxid
                    intMax = GetMaxID("TBLTABLEPROPERTIES", 1, True)
                Else
                    intMax = intMax + 1
                End If

                'first get other instances of this report table 
                idCRT = rowsRT(Count1).Item("ID_TBLCONFIGREPORTTABLES")
                strF2 = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = " & idCRT

                Exit Sub

                Dim rows2() As DataRow = dtblTP.Select(strF2)
                If rows2.Length = 0 Then
                    'big problem
                    var1 = var1
                Else
                    boolHit = True
                    'add a row to dtblTP
                    Dim nr As DataRow = dtblTP.NewRow
                    nr("ID_TBLTABLEPROPERTIES") = intMax
                    For Count2 = 1 To dtblTP.Columns.Count - 1
                        nr(Count2) = rows2(0).Item(Count2)
                    Next Count2
                    dtblTP.Rows.Add(nr)
                End If

            End If

        Next Count1

        If boolHit Then

            If boolGuWuOracle Then
                Try
                    ta_tblTableProperties.Update(tblTableProperties)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLTABLEPROPERTIES.Merge('ds2005.TBLTABLEPROPERTIES, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblTablePropertiesAcc.Update(tblTableProperties)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblTablePropertiesSQLServer.Update(tblTableProperties)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
                End Try
            End If

        End If

end1:

    End Sub

    Sub FillStabilities()

        Dim boolH As Boolean
        boolH = boolHold
        boolHold = True

        Select Case gidCRT
            Case 12, 18, 19, 21, 22, 23, 29, 31, 32 '20190220 LEE: Added 12=Dilution
            Case Else
                Exit Sub
        End Select

        '20181111 LEE
        'Stored in BOOLSTATSNR

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


        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView

        Dim rb1 As System.Windows.Forms.RadioButton
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim var1, var2
        Dim Count1 As Integer

        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)

        If rows.Length = 0 Then
            var1 = -1
        Else
            var1 = NZ(rows(0).Item("BOOLSTATSNR"), -1)
        End If


        'now make the correct choice
        Select Case var1
            Case 1
                rb1 = Me.rbNA
            Case 2
                rb1 = Me.rbProcess
            Case 3
                rb1 = Me.rbBenchTop
            Case 4
                rb1 = Me.rbFT
            Case 5
                rb1 = Me.rbLT
            Case 6
                rb1 = Me.rbReinjection
            Case 7
                rb1 = Me.rbBlood
            Case 8
                rb1 = Me.rbStockSolution
            Case 9
                rb1 = Me.rbSpiking

            Case 10
                rb1 = Me.rbAutosampler
            Case 11
                rb1 = Me.rbBatchReinjection
            Case 12 '20190220 LEE:
                rb1 = Me.rbDilution
            Case Else
                rb1 = Me.rbNA
        End Select

        rb1.Checked = True

end1:

        boolHold = boolH

    End Sub

    Sub ConfigdgvAnalytes(ByVal idR As Int64)

        If boolConfigISR Then
            Exit Sub
        End If

        'add analytes
        Dim var1, var2, var3
        Dim Count1 As Short
        Dim strF As String
        Dim strF1 As String
        Dim rows() As DataRow
        Dim Count2 As Short
        Dim str1 As String
        Dim id1 As Int64
        Dim intG As Short


        strF = "ID_TBLREPORTTABLE = " & idR & " AND ID_TBLSTUDIES = " & id_tblStudies & " AND INTGROUP > 0"

        Try
            'Dim dv As System.Data.DataView = New DataView(tblReportTableAnalytes, strF, "CHARANALYTE", DataViewRowState.CurrentRows)

            Dim dv As System.Data.DataView = New DataView(tblReportTableAnalytes, strF, "ANALYTEID ASC", DataViewRowState.CurrentRows)

            'make a new table and add column "CHARANALYTE" and "MATRIX"
            Dim dtbl As DataTable = dv.ToTable

            str1 = "CHARANALYTE"
            Dim col1 As New DataColumn
            col1.ColumnName = str1
            col1.AllowDBNull = True
            dtbl.Columns.Add(col1)

            str1 = "MATRIX"
            Dim col2 As New DataColumn
            col2.ColumnName = str1
            col2.AllowDBNull = True
            dtbl.Columns.Add(col2)

            'add data
            For Count1 = 0 To dv.Count - 1
                id1 = dv(Count1).Item("ANALYTEID")
                intG = dv(Count1).Item("INTGROUP")
                strF1 = "ANALYTEID = " & id1 & " AND INTGROUP = " & intG
                strF1 = "INTGROUP = " & intG
                
                Try
                    Dim rows1() As DataRow = tblAnalyteGroups.Select(strF1)
                    If rows1.Length = 0 Then
                    Else
                        var2 = rows1(0).Item("ANALYTEDESCRIPTION")
                        var3 = rows1(0).Item("MATRIX")
                        dtbl.Rows(Count1).BeginEdit()
                        dtbl.Rows(Count1).Item("CHARANALYTE") = var2
                        dtbl.Rows(Count1).Item("MATRIX") = var3
                        dtbl.Rows(Count1).EndEdit()
                    End If
                 
                Catch ex As Exception

                End Try

            Next

            Dim dv1 As DataView = New DataView(dtbl, "CHARANALYTE IS NOT NULL", "", DataViewRowState.CurrentRows)

            dv1.AllowDelete = False
            dv1.AllowNew = False

            dv1.Sort = "CHARANALYTE ASC"

            Dim dgv As DataGridView

            dgv = Me.dgvAnalytes

            dgv.DataSource = dv1

            dgv.DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Regular)

            'debug
            For Count1 = 0 To dgv.Columns.Count - 1
                var1 = dgv.Columns(Count1).Name
                var1 = var1
            Next

            If dgv.Columns(1).Visible Then
                For Count1 = 0 To dgv.Columns.Count - 1
                    dgv.Columns(Count1).ReadOnly = True
                    dgv.Columns(Count1).Visible = False
                Next
            End If

            dgv.Columns("CHARANALYTE").HeaderText = "Analyte"
            dgv.Columns("CHARANALYTE").Visible = True
            dgv.Columns("CHARANALYTE").DisplayIndex = 0

            dgv.Columns("MATRIX").HeaderText = "Matrix"
            dgv.Columns("MATRIX").Visible = True
            dgv.Columns("MATRIX").DisplayIndex = 1


            dgv.Columns("NUMINCSAMPLECRIT01").HeaderText = "Acc. Crit. (%)"
            dgv.Columns("NUMINCSAMPLECRIT01").Visible = True
            dgv.Columns("NUMINCSAMPLECRIT01").ReadOnly = False
            dgv.Columns("NUMINCSAMPLECRIT01").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            dgv.RowHeadersWidth = 20
        Catch ex As Exception
            var1 = ex.Message
        End Try


        boolConfigISR = True


    End Sub

    Sub DoThis(ByVal cmd As String)

        Cursor.Current = Cursors.WaitCursor
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short

        boolTS = True

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow
        rows = tblPermissions.Select(strF)

        If StrComp(cmd, "Logoff", CompareMethod.Text) = 0 Then
        Else
            If rows.Length = 0 And boolRefresh = False Then
                MsgBox("Guest does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
                Exit Sub
            End If
        End If

        boolA = BOOLADVANCEDTABLE
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If

        If bool Then
            'Call LockAssignedSamples(bool)
        Else
            MsgBox("This user does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        Select Case cmd
            Case "Edit"

                Call LockWindow(Not (bool))

                'Call DoLegendThings()

                Call SetToEditMode()

                'darken dgvSAS cells
                Call SASReadOnly()


            Case "Save"

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

                If gboolAuditTrail And gboolESig Then

                    Dim frm As New frmESig

                    frm.ShowDialog()

                    If frm.boolCancel Then
                        frm.Dispose()
                        GoTo end1
                    End If

                    gUserID = frm.tUserID
                    gUserName = frm.tUserName

                    frm.Dispose()

                End If

                'clear audittrailtemp
                tblAuditTrailTemp.Clear()
                idSE = 0

                Dim dt1 As DateTime
                dt1 = Now


                Call SaveData()

                '20181112 LEE:
                Call FillTableStuffMethVal(True)

                'record tblaudittrailtemp
                Call RecordAuditTrail(False, dt1)

                Call LockWindow(True)

                Call SetToNonEditMode()

                Call SetTablePropertiesBool(gidTR, gidCRT)

              


            Case "Cancel"

                Call DoCancel(False)
                Call LockWindow(True)
                Call SetToNonEditMode()

                Call FilterSAS()
                'fill values
                Call CallFillSASValues()

                'darken dgvSAS cells
                Call SASReadOnly()

                'check for default settings
                If Me.gbCalcs.Visible Then

                    If Me.rbDifference.Checked Or Me.rbRecovery.Checked Or Me.rbMeanAcc.Checked Then
                    Else
                        'make default
                        Me.rbDifference.Checked = True
                    End If

                End If

                Call UpdateCalc(False)

                Call SetTablePropertiesBool(gidTR, gidCRT)


        End Select

        Call DoMFTable()
        Call UpdateChkData01()
        Call UpdateME01()
        Call DolblMF()
        Call UpdateNomDenom()

        '20181010 LEE:
        'Don't DoLegendThings with DoThis
        'causes changes to be overwritten
        Call DoLegendThings()

        Call DoInjCol()

        '20190109 LEE:
        Call UpdateFDARef()

end1:

        '20181112 LEE:
        Me.lblRemember.Visible = False

        Cursor.Current = Cursors.Default

        boolTS = False

        boolHold = False

    End Sub

    Sub SaveData()

        Dim var1

        Call SaveAutoAssignSamples()

        Call FillAuditTrailTemp(tblTableProperties)

        'temp
        'tblASP.AcceptChanges()

        If boolGuWuOracle Then
            Try
                ta_tblTableProperties.Update(tblTableProperties)
            Catch ex As DBConcurrencyException
                'ds2005.TBLTABLEPROPERTIES.Merge('ds2005.TBLTABLEPROPERTIES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblTablePropertiesAcc.Update(tblTableProperties)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblTablePropertiesSQLServer.Update(tblTableProperties)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
            End Try
        End If

        'also save tblReportTable
        'this will call FillAuditTrailTemp again
        Call SaveTableReportData()

        ''update tblMethodValidation if Method Val Study
        'Try
        '    Dim dgv As DataGridView = frmH.dgvReports
        '    Dim strType As String = NZ(dgv("CHARREPORTTYPE", 0).Value, "Sample Analysis")

        '    If InStr(1, strType, "Validation", CompareMethod.Text) > 0 Then
        '        Call FillTableStuffMethVal(True)
        '    End If
        'Catch ex As Exception

        'End Try


        'also save tblReportTableAnalytes

        Call FillAuditTrailTemp(tblReportTableAnalytes)

        If boolGuWuOracle Then
            Try
                ta_tblReportTableAnalytes.Update(tblReportTableAnalytes)
            Catch ex As DBConcurrencyException
                'ds2005.TBLREPORTTABLEANALYTES.Merge('ds2005.TBLREPORTTABLEANALYTES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblReportTableAnalytesAcc.Update(tblReportTableAnalytes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLREPORTTABLEANALYTES.Merge('ds2005Acc.TBLREPORTTABLEANALYTES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblReportTableAnalytesSQLServer.Update(tblReportTableAnalytes)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLREPORTTABLEANALYTES.Merge('ds2005Acc.TBLREPORTTABLEANALYTES, True)
            End Try
        End If

        Call UpdateTablePropBools()


        '20181112 LEE:
        'must save tblmethodvalidation and tblReports
        Call SaveTableReportData()
        Call SaveMethValTab()


    End Sub

    Sub LockWindow(ByVal bool)

        Dim Count1 As Short

        Me.cmdBuild.Enabled = Not (bool)

        Me.gbAdditional.Enabled = Not (bool)
        Me.gbRTC_Samples.Enabled = Not (bool)
        Me.gbLegendFormat.Enabled = Not (bool)
        Me.gbRTC_CalStd.Enabled = Not (bool)
        Me.gbRTC_QC.Enabled = Not (bool)
        Me.gbStats.Enabled = Not (bool)
        Me.gbAnovaStats.Enabled = Not (bool)

        Me.gbSampleGroup.Enabled = Not (bool)
        Me.gbSampleSort.Enabled = Not (bool)
        Me.gbSAT.Enabled = Not (bool)
        Me.gbPSAE.Enabled = Not (bool)
        Me.gbQCGroup.Enabled = Not (bool)
        Me.gbResultsChoice.Enabled = Not (bool)
        Me.gbTableLegend.Enabled = Not (bool)

        Me.gbIncSampleCriteria.Enabled = Not (bool)

        Me.panNomDenom.Enabled = Not (bool)

        Me.gbCarryover.Enabled = Not (bool)
        Me.gbMatrixFactor.Enabled = Not (bool)

        Me.gbCriteria.Enabled = Not (bool)

        Me.gbRegrULOQ.Enabled = Not (bool)

        Me.dgvAnalytes.ReadOnly = bool
        Try
            Me.dgvSAS.ReadOnly = bool
        Catch ex As Exception

        End Try


        ' Me.dgvAnalytes.Columns("CHARANALYTE").ReadOnly = True
        Try
            Me.dgvAnalytes.Columns("CHARANALYTE").ReadOnly = True
        Catch ex As Exception

        End Try

        Try
            If bool Then
            Else
                Me.dgvAnalytes.Columns("NUMINCSAMPLECRIT01").ReadOnly = False
            End If
        Catch ex As Exception

        End Try

        Try
            'pesky
            For Count1 = 0 To Me.dgvSAS.Columns.Count - 1
                Me.dgvSAS.Columns(Count1).ReadOnly = True
            Next
            If bool Then
            Else
                Me.dgvSAS.Columns("CHARVALUE").ReadOnly = False
                Me.dgvSAS.Columns("CHARNOT").ReadOnly = False
            End If
        Catch ex As Exception

        End Try


        Me.txtTitle.Enabled = Not (bool)

        Me.gbStabilityType.Enabled = Not (bool)

        If bool Then
            Me.cmdIncSamples.Enabled = False
        Else

            Call Disable()

        End If

    End Sub

    Sub DoCancel(ByVal bool)

        Call DoPropCancel() 'this will also cancel Report Table stuff
        Call DoASPCancel()
        Call DoSASCancel()

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        boolFromEdit = True
        Call DoThis("Edit")
        boolFromEdit = False
    End Sub

    Private Sub cmdExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        'Me.Dispose()

        'frmH.dgvReportTableConfiguration.Visible = False

        'Me.Visible = False

        boolClose = True

        Me.lblClose.Left = 0
        Me.lblClose.Top = 0
        Me.lblClose.Height = Me.Height
        Me.lblClose.Width = Me.Width

        Me.lblClose.Visible = True

        If Len(strFilter) = 0 Then
        Else

            Cursor.Current = Cursors.WaitCursor

            Dim dv As DataView = frmH.dgvReportTableConfiguration.DataSource

            dv.RowFilter = strFilter

            'redo colors

            frmH.dgvReportTableConfiguration.AutoResizeRows()

            'frmH.dgvReportTableConfiguration.Visible = True

            'Call ResizeRows(Me.dgvCompanyAnalRef)
            Call AssessSampleAssignment()

            Call UpdateTablePropBools()

            Call AssessQCs()

            Cursor.Current = Cursors.Default

        End If

        Me.Dispose()


    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Call DoThis("Cancel")

    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Call DoThis("Save")

    End Sub

    Private Sub frmReportTableConfig_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '20181218 LEE:
        'Don't know how this happened, but TBLTABLEPROPERTIES isn't in sync with TBLREPORTTABLE in Frontage study BTM-2421
        'a cursory check of other studies doesn't show this
        'will run this code at the opening of Report Table Advanced Configuration just in case
        Call Check_ID_TblReportTable()

        '20180807 LEE:
        boolFormLoad = True

        Cursor.Current = Cursors.WaitCursor

        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

        Dim str1 As String

        str1 = "1. Select table entries to view settings on right." & ChrW(10)
        str1 = str1 & "2. User may manually modify Table Title in text box below." & ChrW(10)
        str1 = str1 & "3. Right-click in Table Title text box below to add Field Codes (if needed)."

        Me.lblTitle.Text = str1

        Me.lblTitle.Top = 5 ' Me.txtTitle.Top - Me.lblTitle.Height + 5

        str1 = "Convert 'deg C' to '" & ChrW(176) & "C'"
        Me.chkCONVERTTEMP.Text = str1

        str1 = ChrW(125)
        Me.lblDiff.Text = str1

        Call SetFormSize()
        Call SetControlSizes()

        Call FormatReportTable()

        Call UpdateCalc(False)

        Call UpdateChkData01()
        Call UpdateMF01()

        Call SetToNonEditMode()

        Me.chkCV.Text = ReturnPrecLabel()

        'pesky
        Call SetC1()
        Call SetC2()

        Cursor.Current = Cursors.Default

        boolFormLoad = False
        boolHold = False

        Call UpdateCalc(False)

        Call DolblMF()

        Call UpdateNomDenom()

        Call ReportTablesChange()

        '20181010 LEE:
        'Don't do this here
        Call DoLegendThings()

        Call DoInjCol()

        boolHold = False

    End Sub


    Sub FormLoad()

        Cursor.Current = Cursors.WaitCursor

        Call MakeSASTable() 'do this temporarily

        Dim Count1 As Short
        Dim dgv As DataGridView

        Dim str1 As String
        str1 = "(1) Accepted characters:  a - z, A - Z, 0 - 9, dash, underscore, period, space, backslash, forward slash, # (no commas)"
        str1 = "(1) Invalid characters: apostrophe (')"
        Me.lblAccepted.Text = str1

        str1 = "(2) Note that Watson sample label is restricted to 20 characters."
        Me.lblWatsonE.Text = str1

        str1 = "(3) All displayed rows are optional."
        Me.lblOptional.Text = str1

        str1 = "(4) Exclude samples that include these terms. Simple OR logic may be used."
        Me.lblLogic.Text = str1

        str1 = "(5) Text Fragments and Exclusions are case-sensitive."
        Me.lblCase.Text = str1

        str1 = "(6) 'Run Identifier' entries: If shown, enter values to add to data set Run ID labels."
        Me.lblRunIdentifier.Text = str1

        boolFormLoad = True

        Call LockWindow(True)
        Call FillAscDesc()
        Call FillGroupSort()

        dgv = Me.dgvReportTables

        If Me.dgvReportTables.RowCount = 0 Then
            Me.cmdEdit.Enabled = False
        Else
            Me.cmdEdit.Enabled = True
        End If

        'Note: dgv datasourse has been set at frmHome.cmdAdvancedTable_Click

        Call FormatReportTableInit()

        'see end of FormLoad for additional dgv formatting

        'Call InsertDefault(-1)

        Me.txtTitle.Text = ""
        If dgv.ColumnCount = 0 Then
        Else

            Dim intRow As Short
            'Dim idSel As Int64
            Dim idSel1 As Int64

            'find introw
            intRow = intORow
            'Try
            '    idSel = frmH.dgvReportTableConfiguration("ID_TBLREPORTTABLE", intORow).Value
            'Catch ex As Exception
            '    idSel = 0
            'End Try
            intRow = 0
            For Count1 = 0 To dgv.Rows.Count - 1
                idSel1 = dgv("ID_TBLREPORTTABLE", Count1).Value
                If idSel = idSel1 Then
                    intRow = Count1
                    Exit For
                End If
            Next

            Try
                boolHold = False
                dgv.CurrentCell = dgv.Rows.Item(intRow).Cells("CHARHEADINGTEXT")
                dgv.FirstDisplayedScrollingRowIndex = intRow
                boolHold = True

            Catch ex As Exception

            End Try

            ''select row intORow
            'If dgv.RowCount = 0 Then 'ignore
            '    intORow = -1
            'ElseIf intORow = -1 Then 'set first row
            '    intORow = 0
            'End If
            'If intORow = -1 Then 'ignore
            'Else 'select first row
            '    dgv.Rows.Item(intORow).Selected = True
            '    boolHold = False
            '    dgv.CurrentCell = dgv.Rows.Item(intORow).Cells("CHARHEADINGTEXT")
            '    boolHold = True
            'End If

        End If

        'fill arrBackup
        Dim var1
        Dim Count2 As Short

        Dim dtbl As DataTable = tblReportTables

        ReDim arrBU(dtbl.Columns.Count, dtbl.Rows.Count)
        'now load rows to arrbu
        For Count1 = 0 To dtbl.Rows.Count - 1
            For Count2 = 0 To dtbl.Columns.Count - 1
                arrBU(Count2 + 1, Count1 + 1) = dtbl.Rows(Count1).Item(Count2)
            Next
        Next

        'set window
        Me.Top = 0
        Dim ht
        'ht = My.Computer.Screen.WorkingArea.Height
        'Me.Height = ht



        Try
            Call frmReportTableConfig_ToolTipSet()
        Catch ex As Exception

        End Try

        'additional dgv formatting

        Call FormatReportTable()


        boolFormLoad = False
        boolHold = False

        Call UpdateCalc(False)

        '20190109 LEE:
        Call UpdateFDARef()

        'for now, delete tab 2
        'Me.tabRTC.TabPages.Remove(tpAutoAssignment)

        Cursor.Current = Cursors.Default

        boolHold = False

    End Sub

    Sub FormatReportTableInit()

        Dim dgv As DataGridView = Me.dgvReportTables
        Dim Count1 As Int16

        dgv = Me.dgvReportTables

        If Me.dgvReportTables.RowCount = 0 Then
            Me.cmdEdit.Enabled = False
        Else
            Me.cmdEdit.Enabled = True
        End If

        'Note: dgv datasourse has been set at frmHome.cmdAdvancedTable_Click

        'hide all columns
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'make two columns visible
        'dgv.Columns("ID_TBLREPORTTABLE").Visible = True'DEBUG

        'dgv.Columns("ID_TBLCONFIGREPORTTABLES").Visible = True'DEBUG

        dgv.Columns("CHARTABLENAME").Visible = True
        dgv.Columns("CHARHEADINGTEXT").Visible = True
        dgv.Columns("CHARFCID").Visible = True

        dgv.Columns("CHARHEADINGTEXT").HeaderText = "Table Title"
        dgv.Columns("CHARTABLENAME").HeaderText = "Table Description"
        dgv.Columns("CHARFCID").HeaderText = "FC ID"

        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        Dim boolF As Boolean = boolFormLoad
        boolFormLoad = True
        dgv.Columns("CHARTABLENAME").ReadOnly = True
        dgv.Columns("CHARHEADINGTEXT").ReadOnly = True
        dgv.Columns("CHARFCID").ReadOnly = True

        '20181128 LEE
        'show the FCID column
        dgv.Columns("CHARHEADINGTEXT").DisplayIndex = 1
        dgv.Columns("CHARFCID").DisplayIndex = 2
        dgv.Columns("CHARTABLENAME").DisplayIndex = 3
        'repeat
        dgv.Columns("CHARHEADINGTEXT").DisplayIndex = 1
        dgv.Columns("CHARFCID").DisplayIndex = 2
        dgv.Columns("CHARTABLENAME").DisplayIndex = 3

        Dim var1
        Try
            dgv.AutoResizeColumn(dgv.Columns("CHARFCID").DisplayIndex)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        boolFormLoad = boolF

        dgv.RowHeadersWidth = 25

    End Sub

    Sub SetFormSize()

        Dim h, w

        h = Screen.PrimaryScreen.WorkingArea.Height
        w = Screen.PrimaryScreen.WorkingArea.Width

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        'taskbar height
        Dim taskbh As Int16 = Screen.PrimaryScreen.Bounds.Height - Screen.PrimaryScreen.WorkingArea.Height
        Dim taskbw As Int16 = Screen.PrimaryScreen.Bounds.Width - Screen.PrimaryScreen.WorkingArea.Width

        Dim h1, w1

        If taskbw = 0 Then
            h1 = h - taskbh ' - 10
            w1 = w - 40

            Me.Location = New Point(10, 10)

            Me.ClientSize = New Size(w1, h1)

        Else 'taskbar is left or right

            h1 = h - taskbh - 10
            w1 = w - taskbw - 40

            Me.Location = New Point(10, 10)

            Me.ClientSize = New Size(w1, h1)

        End If

    

    End Sub


    Sub SetControlSizes()

        Dim t1
        Dim dgv As DataGridView = Me.dgvReportTables

        Me.txtTitle.Left = 12
        Me.txtTitle.Top = 58
        Me.txtTitle.Width = 600

        t1 = Me.txtTitle.Top

        dgv.Left = Me.txtTitle.Left
        dgv.Top = t1 + Me.txtTitle.Height + 10
        dgv.Width = Me.txtTitle.Width
        dgv.Height = Me.ClientSize.Height - dgv.Top - 15

        'Me.tabRTC.Left = Me.txtTitle.Left + Me.txtTitle.Width + 10
        'Me.tabRTC.Width = Me.gbAdditional.Width + 15 + 10 + Me.panFormat.Left + 15

        Call SetC1()
        Call SetC2()

        'make all group boxes the same width as gbTableLegend
        Dim intW As Int64 = gbTableLegend.Width

        Dim ctrl As Control
        Dim str1 As String
        Dim str2 As String

        For Each ctrl In Me.panFormat.Controls
            str1 = ctrl.Name
            str2 = Mid(str1, 1, 2)
            If StrComp(str2, "gb", CompareMethod.Text) = 0 And StrComp(str1, "gbMatrixFactor", CompareMethod.Text) <> 0 Then
                ctrl.Width = intW
            End If
        Next

        'cover pbxTableGraphicExamples
        Dim a, b, c

        a = Me.panTableGraphicExamples.Left + Me.panTableGraphicExamples.Width
        b = a - Me.dgvSAS.Left
        Me.dgvSAS.Width = b


    End Sub

    Sub SetC1()

        Dim t1
        Dim dgv As DataGridView = Me.dgvReportTables

        t1 = Me.txtTitle.Top

        Me.tabRTC.Top = t1
        Me.tabRTC.Left = 622 ' Me.txtTitle.Left + Me.txtTitle.Width + 10
        Me.tabRTC.Height = Me.ClientSize.Height - Me.tabRTC.Top - 15

        Me.panEdit.Left = Me.tabRTC.Left
        Me.panEdit.Top = 5 'Me.tabRTC.Top - Me.panEdit.Height - 2

        Me.cmdSymbol.Top = Me.tabRTC.Top - 15


    End Sub

    Sub SetC2()

        Dim x, y, z

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        x = Me.tabRTC.Left
        y = Me.ClientSize.Width - bw

        Me.tabRTC.Width = y - x - 10

        Me.cmdSymbol.Left = Me.tabRTC.Left + Me.tabRTC.Width - Me.cmdSymbol.Width


    End Sub



    Sub FillAscDesc()

        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim c As Control

        For Count1 = 1 To 10
            Select Case Count1
                Case 1, 2, 3, 4
                    str1 = "cbxSampleGAD" & Count1
                    For Each c In Me.gbSampleGroup.Controls
                        str2 = c.Name
                        If StrComp(str2, str1, CompareMethod.Text) = 0 Then
                            Dim cbx As ComboBox = c
                            cbx.Items.Add("ASC")
                            cbx.Items.Add("DESC")
                            cbx.SelectedIndex = 0
                            Exit For
                        End If
                    Next

                    'Case 5, 6, 7, 8
                Case Else
                    str1 = "cbxSampleSAD" & Count1 - 4
                    For Each c In Me.gbSampleSort.Controls
                        str2 = c.Name
                        If StrComp(str2, str1, CompareMethod.Text) = 0 Then
                            Dim cbx As ComboBox = c
                            cbx.Items.Add("ASC")
                            cbx.Items.Add("DESC")
                            cbx.SelectedIndex = 0
                            Exit For
                        End If
                    Next

            End Select

        Next

    End Sub

    Sub FillGroupSort()

        Dim c As Control
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intRows As Short

        'dtbl = tblReportTableHeaderConfig
        'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLCONFIGREPORTTABLES = 5"

        dtbl = tblConfigHeaderLookup
        strF = "ID_TBLCONFIGREPORTTABLES = 5"

        strS = "INTORDER ASC"
        rows = dtbl.Select(strF, strS)
        intRows = rows.Length

        For Count1 = 1 To 10
            Select Case Count1
                Case 1, 2, 3, 4
                    str1 = "cbxSampleG" & Count1
                    For Each c In Me.gbSampleGroup.Controls
                        str2 = c.Name
                        If StrComp(str2, str1, CompareMethod.Text) = 0 Then
                            Dim cbx As ComboBox = c
                            cbx.Items.Add("[None]")
                            For Count2 = 0 To intRows - 1
                                str3 = rows(Count2).Item("CHARCOLUMNLABEL")
                                cbx.Items.Add(str3)
                            Next
                            cbx.SelectedIndex = 0
                            Exit For
                        End If
                    Next
                    'Case 5, 6, 7, 8
                Case Else
                    str1 = "cbxSampleS" & Count1 - 4
                    For Each c In Me.gbSampleSort.Controls
                        str2 = c.Name
                        If StrComp(str2, str1, CompareMethod.Text) = 0 Then
                            Dim cbx As ComboBox = c
                            cbx.Items.Add("[None]")
                            For Count2 = 0 To intRows - 1
                                str3 = rows(Count2).Item("CHARCOLUMNLABEL")
                                'If StrComp(str3, "Group", CompareMethod.Text) = 0 Then
                                '    str3 = str3 & " Name"
                                'ElseIf StrComp(str3, "Subject", CompareMethod.Text) = 0 Then
                                '    str3 = str3 & " Name"
                                'End If
                                cbx.Items.Add(str3)
                                'If InStr(1, str3, "Group", CompareMethod.Text) > 0 Then
                                '    str3 = "Group ID"
                                '    cbx.Items.Add(str3)
                                'ElseIf InStr(1, str3, "Subject", CompareMethod.Text) > 0 Then
                                '    str3 = "Subject ID"
                                '    cbx.Items.Add(str3)
                                'End If
                            Next

                            '20170405 LEE: Need to allow users to choose:
                            '  - Group Name or ID
                            '  - Subject Name or ID
                            'str3 = "Group ID"

                            cbx.SelectedIndex = 0
                            Exit For
                        End If
                    Next
            End Select

        Next

    End Sub


    Sub FormatReportTable()

        Try
            Dim dgv As DataGridView = Me.dgvReportTables

            'Dim newPadding As New Padding(0, 1, 0, CUSTOM_CONTENT_HEIGHT)
            'Me.dataGridView1.RowTemplate.DefaultCellStyle.Padding = newPadding
            'Left, Top, Right, and Bottom properties,

            'dgv.Columns("CHARHEADINGTEXT").HeaderText = "Table Title"
            'dgv.Columns("CHARTABLENAME").HeaderText = "Table Description"

            Dim nP As New Padding(0, 6, 0, 6)
            dgv.DefaultCellStyle.Padding = nP

            dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

            dgv.Columns("CHARHEADINGTEXT").MinimumWidth = dgv.Width * 0.75
            dgv.Columns("CHARHEADINGTEXT").Width = dgv.Width * 0.75

            'dgv.AutoResizeColumn("CHARTABLENAME")

            Dim int1 As Short
            int1 = dgv.Columns("CHARTABLENAME").Index
            dgv.AutoResizeColumn(int1)

            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dgv.AutoResizeRows()

            dgv.Columns("CHARHEADINGTEXT").MinimumWidth = dgv.Width * 0.75
            dgv.Columns("CHARHEADINGTEXT").Width = dgv.Width * 0.75

            'do same for other tables
            Dim nP1 As New Padding(0, 3, 0, 3)
            Me.dgvSAS.DefaultCellStyle.Padding = nP1

        Catch ex As Exception

        End Try


    End Sub

    Private Sub dgvReportTables_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportTables.CellContentClick

    End Sub

    Private Sub dgvReportTables_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportTables.SelectionChanged

        If boolHold Then
            'Exit Sub
        End If

        If boolFormLoad Then
            Exit Sub
        End If

        If boolClose Then
            Exit Sub
        End If

        Call ReportTablesChange()


    End Sub

    Sub ReportTablesChange()

        boolTS = True

        Call FillProperties()
        Call Disable()
        Call FilterSAS()



        'Call DoLegendThings()


        'check for default settings
        If Me.gbCalcs.Visible Then

            If Me.rbDifference.Checked Or Me.rbRecovery.Checked Or Me.rbMeanAcc.Checked Then
            Else
                'make default
                Me.rbDifference.Checked = True
            End If

        End If

        Call DoMFTable()

        Call UpdateME01()
        Call UpdateChkData01()
        Call UpdateMF01()
        Call DolblMF()

        Call UpdateNomDenom()

        '20181010 LEE: Don't do this here
        Call DoLegendThings()

        Call DoInjCol()

        '20190109 LEE:
        Call UpdateFDARef()

        boolTS = False

    End Sub


    Sub Disable()

        Dim dgv As DataGridView
        Dim tblP As System.Data.DataTable
        Dim rowsP() As DataRow
        Dim tblC As System.Data.DataTable
        Dim rowsC() As DataRow
        Dim intRow As Short
        Dim idT As Int64
        Dim idC As Int64
        Dim strF As String
        Dim var1
        Dim Count1 As Short
        Dim str1 As String
        Dim bool As Boolean
        Dim boolCS As Boolean
        Dim boolQC As Boolean
        Dim boolRC As Boolean 'gbResultsChoice
        Dim intAssign As Short

        Dim gbIncr As Short = 2

        Dim t1, l1, t2

        'idC legend
        '1: Summary of Analytical Runs
        '2: Summary of Regression Constants
        '3: Summary of Back-Calculated Calibration Std Conc
        '4: Summary of Interpolated QC Std Conc
        '5: Summary of Samples
        '6: Summary of Reassayed Samples
        '7: Summary of Repeat Samples
        '11: Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
        '12: Summary of Interpolated Dilution QC Concentrations
        '13: Summary of Combined Recovery
        '14: Summary of True Recovery
        '15: Summary of Suppression/Enhancement
        '17: Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
        '18: Summary of [Period Temp] Stability in Matrix
        '19: Summary of Freeze/Thaw [#Cycles] Stability in Matrix
        '21: [Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
        '22: [Period Temp] Stock Solution Stability Assessment
        '23: [Period Temp] Spiking Solution Stability Assessment
        '29: [Period Temp] Long-Term QC Std Storage Stability
        '30: Incurred Samples
        '31: Ad Hoc QC Stability Table
        '32: Ad Hoc QC Stability Comparison Table


        tblP = tblTableProperties
        tblC = tblConfigReportTables

        Dim i As Short
        i = 5
        boolCS = False
        boolQC = False

        dgv = Me.dgvReportTables

        t1 = dgv.Top
        t1 = Me.txtTitle.Top
        l1 = dgv.Left + dgv.Width + i

        l1 = 5

        'Me.cmdIncSamples.Top = t1
        'Me.cmdIncSamples.Left = l1
        't2 = Me.cmdIncSamples.Height + gbIncr
        't1 = t1 + t2

        'Me.gbIncSampleCriteria.Top = t1
        'Me.gbIncSampleCriteria.Left = l1
        't2 = Me.gbIncSampleCriteria.Height + gbIncr
        't1 = t1 + t2


        t1 = Me.txtTitle.Top

        'These numbers were derived empirically (NDL)
        '2015113 LEE: moved these to SetControlSizes
        'Me.tabRTC.Width = Me.gbAdditional.Width + 15 + 10 + Me.panFormat.Left + 15
        'Me.tabRTC.Top = t1
        'Me.tabRTC.Left = Me.txtTitle.Left + Me.txtTitle.Width + 2

        'dgv.Height = Me.ClientSize.Height - dgv.Top - 15
        'Me.tabRTC.Height = Me.ClientSize.Height - Me.tabRTC.Top - 15

        l1 = Me.tabRTC.Left + Me.tabRTC.Width

        'Me.cmdCopySymbol.Left = l1 + 2
        'Me.lblSymbol.Left = l1 + 2
        'Me.lblSymbol1.Left = l1 + 2
        'Me.lbxSymbol.Left = l1 + 2
        'Me.txtSymbol.Left = l1 + 2

        'Me.Width = Me.lblSymbol1.Left + Me.lblSymbol1.Width + 25

        l1 = 2.5

        t1 = 5

        If dgv.RowCount = 0 Then
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        intAssign = dgv("BOOLREQUIRESSAMPLEASSIGNMENT", intRow).Value

        idC = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        strF = "ID_TBLCONFIGREPORTTABLES = " & idC

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        Select Case idC
            Case 30
                Call ConfigdgvAnalytes(idT)

        End Select


        rowsC = tblC.Select(strF)
        var1 = rowsC(0).Item("BOOLSAMPLES")
        If var1 = -1 Then
            bool = True

            Me.gbRTC_Samples.Top = t1
            Me.gbRTC_Samples.Left = l1
            t2 = Me.gbRTC_Samples.Height + gbIncr
            t1 = t1 + t2

            If idC = 6 Then
                Me.gbLegendFormat.Top = t1
                Me.gbLegendFormat.Left = l1
            End If

            If idC = 30 Then
            Else
                Me.gbSampleGroup.Top = t1
                Me.gbSampleGroup.Left = l1
                t2 = Me.gbSampleGroup.Height + gbIncr
                t1 = t1 + t2

                Me.gbSampleSort.Top = t1
                Me.gbSampleSort.Left = l1
                t2 = Me.gbSampleSort.Height + gbIncr
                t1 = t1 + t2

                Me.gbSAT.Top = t1
                Me.gbSAT.Left = l1
            End If

        Else
            bool = False
        End If


        'If idC = 6 Then
        '    Me.gbRTC_Samples.Visible = False
        'Else
        '    Me.gbRTC_Samples.Visible = bool
        'End If

        Me.gbRTC_Samples.Visible = bool

        Select Case idC
            Case 6
                Me.gbLegendFormat.Visible = True
            Case Else
                Me.gbLegendFormat.Visible = False
        End Select

        Select Case idC
            Case 6, 7, 30
                Me.gbSampleGroup.Visible = False
                Me.gbSampleSort.Visible = False
            Case Else
                Me.gbSampleGroup.Visible = bool
                Me.gbSampleSort.Visible = bool
        End Select

        Select Case idC
            Case 5
                Me.gbSAT.Visible = True
            Case Else
                Me.gbSAT.Visible = False
        End Select


        var1 = rowsC(0).Item("BOOLCSSTATS")
        If var1 = -1 Then
            bool = True
            Me.gbRTC_CalStd.Top = t1
            Me.gbRTC_CalStd.Left = l1
            t2 = Me.gbRTC_CalStd.Height + gbIncr
            t1 = t1 + t2

        Else
            bool = False
        End If
        Me.gbRTC_CalStd.Visible = bool
        boolCS = bool

        Select Case idC
            Case 33, 34, 35, 36, 37, 38
                boolCS = False
        End Select

        var1 = rowsC(0).Item("BOOLQCSTATS")
        If var1 = -1 Then
            Select Case idC
                Case 33, 34, 35
                    bool = True
                Case Else
                    bool = True
                    Me.gbRTC_QC.Top = t1
                    Me.gbRTC_QC.Left = l1
                    t2 = Me.gbRTC_QC.Height + gbIncr
                    t1 = t1 + t2

            End Select

        Else
            bool = False
        End If
        Select Case idC
            Case 33, 34, 35, 36, 37, 38
                Me.gbRTC_QC.Visible = False
            Case Else
                Me.gbRTC_QC.Visible = bool
        End Select

        boolQC = bool

        Select Case idC
            '20160905 LEE: Don't show gbResultsChoice for _17: Unique Low QCs
            'This was shown for Matrix Effect features, which no one seems to use in this context
            'See comments in Sub ..._17
            'Case 13, 14, 15, 17, 22, 23, 31, 32
            Case 13, 14, 15, 22, 23, 29, 31, 32
                boolRC = True
            Case Else
                boolRC = False
        End Select

        If boolCS Or boolQC Or idC = 2 Then 'regression table can have stats
            Me.gbStats.Visible = True
            Me.gbStats.Top = t1
            Me.gbStats.Left = l1
            t2 = Me.gbStats.Height + gbIncr
            t1 = t1 + t2

            Me.chkIncludeIS_Single.Top = Me.chkIncludeWatsonLabel.Top
            Me.chkIncludeIS_Single.Left = Me.chkIncludeWatsonLabel.Left

            Select Case idC
                Case 13, 14, 15
                    Me.chkDiffCol.Visible = False
                Case Else
                    Me.chkDiffCol.Visible = True
            End Select

        Else
            Me.gbStats.Visible = False
        End If

        '20181110 LEE:
        Me.gbStabilityType.Left = Me.gbAdditional.Left + Me.gbAdditional.Width + gbIncr
        'make visible for certain tables
        Select Case idC
            Case 12, 18, 19, 21, 22, 23, 29, 31, 32
                Me.gbStabilityType.Visible = True
            Case Else
                Me.gbStabilityType.Visible = False
        End Select
        '20190109 LEE:
        '20190220 LEE:
        Select Case idC
            Case 18, 19, 21, 22, 23, 29, 31, 32 'do not show for 12=Dilution
                'Me.panFDARef.Visible = True
                If idC = 31 Then 'check in case assay is dilution
                    Call StabilityViews()
                Else
                    Me.panFDARef.Visible = True
                End If
            Case Else
                Me.panFDARef.Visible = False
        End Select
        'Me.panFDARef.Visible = Me.gbStabilityType.Visible

        Me.gbCriteria.Top = t1
        Me.gbCriteria.Left = Me.gbStats.Left

        Me.gbRegrULOQ.Top = t1
        Me.gbRegrULOQ.Left = Me.gbStats.Left


        'If idC = 11 Then
        '    Me.gbAnovaStats.Visible = True
        '    Me.gbAnovaStats.Top = t1
        '    Me.gbAnovaStats.Left = l1
        '    t2 = Me.gbAnovaStats.Height + gbIncr
        '    t1 = t1 + t2

        'Else
        '    Me.gbAnovaStats.Visible = False
        'End If


        Select Case idC
            Case 35 'carryover
                Me.gbCarryover.Visible = True
                Me.gbCarryover.Top = t1
                Me.gbCarryover.Left = l1
                t2 = Me.gbCarryover.Height + gbIncr
                t1 = t1 + t2
            Case Else
                Me.gbCarryover.Visible = False
        End Select

        Select Case idC
            Case 11, 13, 14, 15
                Select Case idC
                    Case 11
                        'Me.chkIncludeAnovaSumStats.Top = 38
                        Me.chkIncludeAnova.Visible = True
                        Me.chkIntraRunSumStats.Visible = True
                        Me.chkIncludeAnovaSumStats.Visible = True
                    Case Else
                        'Me.chkIncludeAnovaSumStats.Top = 19
                        Me.chkIntraRunSumStats.Visible = True
                        Me.chkIncludeAnova.Visible = False
                        Me.chkIncludeAnovaSumStats.Visible = False
                End Select
                Me.gbAnovaStats.Visible = True
                Me.gbAnovaStats.Top = t1
                Me.gbAnovaStats.Left = l1
                t2 = Me.gbAnovaStats.Height + gbIncr
                t1 = t1 + t2
            Case Else
                Me.gbAnovaStats.Visible = False
        End Select

        Select Case idC
            Case 35
                Me.chkIncludeIS_Single.Visible = True
            Case Else
                Me.chkIncludeIS_Single.Visible = False
        End Select

        Select Case idC
            Case 2
                Me.gbRegrULOQ.Visible = True
            Case 17
                Me.gbCriteria.Visible = True

            Case Else
                Me.gbRegrULOQ.Visible = False
                Me.gbCriteria.Visible = False

        End Select

        Select Case idC
            Case 22
                Me.chkBOOLISCOMBINELEVELS.Visible = False
                Me.chkCustomLeg.Visible = True
            Case Else
                Me.chkBOOLISCOMBINELEVELS.Visible = True
                Me.chkCustomLeg.Visible = False
        End Select

        '*****
        Me.chkNoneLeg.Visible = True
        Me.gbCalcs.Visible = False

        Me.panIS.Enabled = True 'why is panIS disabled?
        Me.chkIncludeIS.Enabled = True 'why is this disabled?
        Me.chkIncludeIS_Single.Enabled = True
        Me.chkCustomLeg.Enabled = True

        Me.gbMatrixFactor.Visible = False

        Select Case idC
            Case 17
                Me.gbResultsChoice.Visible = True
                Me.gbResultsChoice.Top = t1
                Me.gbResultsChoice.Left = l1
        End Select

        If boolRC Then
            Me.gbResultsChoice.Visible = True
            Me.gbResultsChoice.Top = t1
            Me.gbResultsChoice.Left = l1

            Select Case idC
                Case 15
                    Me.gbMatrixFactor.Visible = True
                    Me.gbMatrixFactor.Top = t1
                    Me.gbMatrixFactor.Left = l1 + Me.gbResultsChoice.Width + 5
                    'Me.gbMatrixFactor.Height = Me.gbResultsChoice.Height
            End Select

            t2 = Me.gbResultsChoice.Height + gbIncr
            t1 = t1 + t2

            '20181130 LEE:
            Me.panNomDenomCalcs.Visible = True
            Me.gbNumerator.Visible = True
            Me.gbDenom.Visible = True

            Me.panNomDenomCalcs.Top = Me.chkNoneLeg.Top - 10
            Me.panTitleLegends.Top = Me.panNomDenomCalcs.Top + Me.panNomDenomCalcs.Height + gbIncr

            Select Case idC
                Case 13, 14, 15
                    Me.rbConc.Visible = False

                    Me.panIS.Visible = True
                    Me.gbCalcs.Visible = False

                    '20181130 LEE:
                    Me.panNomDenomCalcs.Visible = False
                    Me.panTitleLegends.Top = Me.chkNoneLeg.Top + Me.chkNoneLeg.Height + gbIncr

                Case 17
                    Me.rbConc.Visible = True
                    Me.panIS.Visible = True
                    Me.panIS.Enabled = True

                Case 22
                    Me.rbConc.Visible = False
                    Me.panIS.Visible = True
                    Me.gbCalcs.Visible = True

                Case 23
                    Me.rbConc.Visible = False
                    Me.panIS.Visible = True 'False
                    Me.gbCalcs.Visible = True
                Case 31
                    Me.rbConc.Visible = True
                    Me.panIS.Visible = False

                    Me.gbCalcs.Visible = False
                Case 32
                    Me.rbConc.Visible = True

                    Me.panIS.Visible = False

                    Me.gbCalcs.Visible = True
                Case Else
                    Me.rbConc.Visible = True

                    Me.panIS.Visible = False
            End Select

            Select Case idC
                Case 13, 14, 15, 17, 30
                    Me.gbTableLegend.Visible = True
                    Me.panTitleLegends.Visible = True
                Case 22, 23, 29, 32
                    Me.gbTableLegend.Visible = True
                    Me.panTitleLegends.Visible = True
                    Me.gbCalcs.Visible = True
                Case Else
                    Me.gbTableLegend.Visible = False
            End Select

            Select Case idC
                Case 13, 14, 15, 22, 23, 29, 32

                    Me.rbPosLeg.Visible = True
                    Me.rbNegLeg.Visible = True
                    Me.rbMeanAcc.Visible = True
                Case Else

                    Me.rbPosLeg.Visible = False
                    Me.rbNegLeg.Visible = False
                    Me.rbMeanAcc.Visible = False
            End Select

        Else

            Me.rbPosLeg.Visible = False
            Me.rbNegLeg.Visible = False
            Me.gbResultsChoice.Visible = False
            'Me.gbTableLegend.Visible = False

            Select Case idC
                Case 29
                    Me.gbTableLegend.Visible = True
                    Me.panTitleLegends.Visible = False

                Case 13, 14, 15, 30, 34, 35 'carryover 20190108 LEE
                    Me.gbTableLegend.Visible = True
                    Me.panTitleLegends.Visible = False
                    Me.gbCalcs.Visible = False

                Case Else
                    Me.gbTableLegend.Visible = False
            End Select

        End If

        Select Case idC
            Case 32
                Me.chkBOOLADHOCSTABCOMPCOLUMNS.Visible = True
                Me.chkBOOLADHOCSTABCOMPCOLUMNS.Top = Me.rbUsePeakAreaRatio.Top + Me.rbUsePeakAreaRatio.Height + gbIncr
                Me.chkBOOLADHOCSTABCOMPCOLUMNS.Left = Me.rbUsePeakAreaRatio.Left
            Case Else
                Me.chkBOOLADHOCSTABCOMPCOLUMNS.Visible = False
        End Select

        Select Case idC
            Case 17

                'move below gbCriteria
                t2 = Me.gbCriteria.Height + gbIncr
                t1 = t1 + t2
                Me.gbResultsChoice.Top = t1

                Me.gbResultsChoice.Visible = True

            Case Else
                If Me.gbTableLegend.Visible Then
                    Me.gbTableLegend.Visible = True
                    Me.gbTableLegend.Top = t1
                    Me.gbTableLegend.Left = l1
                    t2 = Me.gbTableLegend.Height + gbIncr
                    t1 = t1 + t2

                End If
        End Select

       

        '*****

        var1 = rowsC(0).Item("BOOLREQUIRESPERTEMP")
        If var1 = -1 Then
            bool = True
            If idC = 19 Then
                Me.panCycles.Visible = True
                Me.panTP.Visible = True ' False
            Else
                Me.panCycles.Visible = False
                Me.panTP.Visible = True
            End If

            Me.gbAdditional.Left = 3


        Else
            bool = False
            Me.CHARTIMEPERIOD.Text = ""
            Me.CHARTIMEFRAME.Text = ""
            Me.CHARPERIODTEMP.Text = ""
            Me.INTNUMBEROFCYCLES.Text = ""
            Me.CHARSTABILITYPERIOD.Text = ""
        End If
        Me.gbAdditional.Visible = bool


        'further define gbadditional

        Dim boolI As Boolean = False

        Me.cmdIncSamples.Visible = False

        Select Case idC
            Case 30
                Me.gbIncSampleCriteria.Visible = True
            Case Else
                Me.gbIncSampleCriteria.Visible = False
        End Select

        If boolFromEdit Then
            If idC = 30 Then

                Me.gbIncSampleCriteria.Enabled = True
                boolI = True
            Else
                Me.gbIncSampleCriteria.Visible = False
            End If
        Else
            If Me.cmdEdit.Enabled Then 'ignore
                If idC = 30 Then

                    Me.gbIncSampleCriteria.Enabled = False
                    boolI = True
                Else
                    Me.gbIncSampleCriteria.Visible = False
                End If
            Else
                If idC = 30 Then

                    Me.gbIncSampleCriteria.Enabled = True
                    boolI = True
                Else
                    Me.gbIncSampleCriteria.Visible = False
                End If
            End If
        End If

        If boolI Then
            Me.gbIncSampleCriteria.Top = t1
            Me.gbIncSampleCriteria.Left = l1
            t2 = Me.gbIncSampleCriteria.Height + gbIncr
            t1 = t1 + t2
        End If

        '20180809 LEE:
        'Establish baseline
        Me.gbPSAE.Text = "Include PSAE Data"
        Me.chkIncludePSAE.Text = "Include PSAE Samples"
        Me.chkInjCol.Visible = False

        Me.gbQCGroup.Visible = False
        Me.gbPSAE.Visible = False

        '20180823 LEE:
        'str1 = "Group QC Levels By (Ignored if Assigned Samples):"
        str1 = "Group QC Levels By:"

        'If idC = 1 Or idC = 2 Then'don't do Sum of Anal Runs for now
        Select Case idC
            Case 2

                Me.gbPSAE.Visible = True
                Me.gbPSAE.Top = t1
                Me.gbPSAE.Left = l1
                t2 = Me.gbPSAE.Height + gbIncr
                t1 = t1 + t2

            Case 3, 4

                'ensure that assigned samples is false

                If intAssign = -1 Then
                    Me.gbPSAE.Visible = False
                Else
                    Me.gbPSAE.Visible = True
                    Me.gbPSAE.Top = t1
                    Me.gbPSAE.Left = l1
                    t2 = Me.gbPSAE.Height + gbIncr
                    t1 = t1 + t2
                End If

                If idC = 4 Then
                    Me.gbQCGroup.Text = str1
                    Me.gbQCGroup.Visible = True
                    Me.gbQCGroup.Top = t1
                    Me.gbQCGroup.Left = l1
                    t2 = Me.gbQCGroup.Height + gbIncr
                    t1 = t1 + t2
                End If

            Case 31

                'move these to the left even with previous
                str1 = "Group QC Levels By:"
                Me.gbQCGroup.Text = str1
                Me.gbQCGroup.Visible = True
                Me.gbQCGroup.Top = t1 - t2
                Me.gbQCGroup.Left = l1 + Me.gbTableLegend.Width + gbIncr
                t2 = Me.gbQCGroup.Height + gbIncr
                t1 = t1 + t2

            Case 35

                Me.gbPSAE.Text = "Table Column Configuration"
                Me.chkIncludePSAE.Text = "Exclude ULOQ column"
                Me.chkInjCol.Visible = True

                '20190108 LEE:
                'shrink gbTableLegend
                Me.gbTableLegend.Height = 53
                Me.chkRTC_CalStd_Acc.Visible = False
                Me.panNomDenom.Visible = False
                Me.gbCalcs.Visible = False
                Me.panTitleLegends.Visible = False
                'modify t1
                t1 = t1 - (328 - 53)

                str1 = "Group QC Levels By:"
                Me.gbQCGroup.Text = str1
                Me.gbPSAE.Visible = True
                Me.gbPSAE.Top = t1
                Me.gbPSAE.Left = l1
                t2 = Me.gbPSAE.Height + gbIncr
                t1 = t1 + t2

            Case Else

                Me.gbPSAE.Visible = False
                Me.gbQCGroup.Visible = False
                str1 = "Group QC Levels By (Ignored if Assigned Samples):"
                Me.gbQCGroup.Text = str1

                Me.gbPSAE.Text = "Include PSAE Data"
                Me.chkIncludePSAE.Text = "Include PSAE Samples"

                '20190108 LEE:
                'reset gbTableLegend
                Me.gbTableLegend.Height = 328
                Me.chkRTC_CalStd_Acc.Visible = True
                Me.panNomDenom.Visible = True
                Me.gbCalcs.Visible = True
                Me.panTitleLegends.Visible = True

                '20190108 LEE
                Select Case idC

                    Case 13, 14, 15, 30, 34
                        '20190108 LEE:
                        'shrink gbTableLegend
                        Me.gbTableLegend.Height = 53
                        Me.chkRTC_CalStd_Acc.Visible = False
                        Me.panNomDenom.Visible = False
                        Me.gbCalcs.Visible = False
                        Me.panTitleLegends.Visible = False
                        'modify t1
                        't1 = t1 - (328 - 53)
                        t1 = Me.gbTableLegend.Top + Me.gbTableLegend.Height + gbIncr

                        Select Case idC '20190108 LEE:
                            Case 13, 14, 15
                                'reposition pantitlelegends
                                Me.panTitleLegends.Top = t1
                            Case 30 'ISR, must move gbIncSampleCriteria
                                Me.gbIncSampleCriteria.Top = t1
                                Me.gbIncSampleCriteria.Left = l1
                                t2 = Me.gbIncSampleCriteria.Height + gbIncr
                                t1 = t1 + t2
                        End Select

                End Select

        End Select

        'If idC = 2 Then
        '    Me.gbPSAE.Visible = True
        '    Me.gbPSAE.Top = t1
        '    Me.gbPSAE.Left = l1
        '    t2 = Me.gbPSAE.Height + gbIncr
        '    t1 = t1 + t2
        'ElseIf idC = 3 Or idC = 4 Then
        '    'ensure that assigned samples is false
        '    If intAssign = -1 Then
        '        Me.gbPSAE.Visible = False
        '    Else
        '        Me.gbPSAE.Visible = True
        '        Me.gbPSAE.Top = t1
        '        Me.gbPSAE.Left = l1
        '        t2 = Me.gbPSAE.Height + gbIncr
        '        t1 = t1 + t2
        '    End If

        '    If idC = 4 Then
        '        Me.gbQCGroup.Visible = True
        '        Me.gbQCGroup.Top = t1
        '        Me.gbQCGroup.Left = l1
        '        't2 = Me.gbPSAE.Height + gbIncr
        '        t1 = t1 + t2
        '    Else

        '    End If

        'ElseIf idC = 35 Then 'carryover

        '    Me.gbPSAE.Text = "Table Column Configuration"
        '    Me.chkIncludePSAE.Text = "Exclude ULOQ column"
        '    Me.chkInjCol.Visible = True

        '    Me.gbPSAE.Visible = True
        '    Me.gbPSAE.Top = t1
        '    Me.gbPSAE.Left = l1
        '    t2 = Me.gbPSAE.Height + gbIncr
        '    t1 = t1 + t2

        'Else
        '    Me.gbPSAE.Visible = False
        '    Me.gbQCGroup.Visible = False

        '    Me.gbPSAE.Text = "Include PSAE Data"
        '    Me.chkIncludePSAE.Text = "Include PSAE Samples"

        'End If

        If idC = 3 Then
            Me.chkIncludeWatsonLabel.Visible = True
        Else
            Me.chkIncludeWatsonLabel.Visible = False
        End If

        Select Case idC
            Case 13, 14, 15, 17
                Select Case idC
                    Case 15

                        If BOOLMFTABLE Then
                            str1 = "Calculate individual" & ChrW(10) & "Matrix Factor values"
                        Else
                            str1 = "Calculate individual" & ChrW(10) & "Sup/Enh values"
                        End If
                    Case 17
                        Me.chkMFTable.Visible = True
                        str1 = "Calculate Matrix Effect and Matrix Factor values"
                    Case Else
                        str1 = "Calculate individual" & ChrW(10) & "Recovery values"
                End Select
                Me.chkBOOLDOINDREC.Text = str1
                Me.chkBOOLDOINDREC.Visible = True
            Case Else
                Me.chkBOOLDOINDREC.Visible = False
        End Select


        'gbGroupSort

        Select Case idC
            Case 6
                t1 = Me.gbLegendFormat.Top
                t2 = Me.gbLegendFormat.Height + gbIncr
                t1 = t1 + t2
                Me.gbGroupSort.Top = t1
                Me.gbGroupSort.Left = l1
            Case 7
                t1 = Me.gbRTC_Samples.Top
                t2 = Me.gbRTC_Samples.Height + gbIncr
                t1 = t1 + t2
                Me.gbGroupSort.Top = t1
                Me.gbGroupSort.Left = l1
            Case 30
                t1 = Me.gbIncSampleCriteria.Top
                t2 = Me.gbIncSampleCriteria.Height + gbIncr
                t1 = t1 + t2
                Me.gbGroupSort.Top = t1
                Me.gbGroupSort.Left = l1
        End Select

        Select Case idC
            Case 6, 7, 30
                Me.gbGroupSort.Visible = True
            Case Else
                Me.gbGroupSort.Visible = False
        End Select

        'now do Criteria

        t1 = 5


        t1 = t1 + t2

        Select Case idC
            Case 36, 37, 38
                Me.gbStats.Visible = False
        End Select

        Dim boolC As Boolean = boolTS
        boolTS = True
        Call PARChange(False)
        boolTS = boolC

end1:

    End Sub


    Sub InsertDefault(ByVal intRow)

        Dim dgv As DataGridView
        Dim tblP As System.Data.DataTable
        Dim rowsP() As DataRow
        Dim tblIS As System.Data.DataTable
        Dim rowsIS() As DataRow
        'Dim intRow As Short
        Dim idT As Int64
        Dim idS As Int64
        Dim idC As Int64
        Dim id As Int64
        Dim Count1 As Short

        Dim strF As String
        Dim var1
        Dim str1 As String
        Dim intS As Short
        Dim intE As Short

        tblP = tblTableProperties
        tblIS = tblTableLegends

        dgv = Me.dgvReportTables
        If dgv.RowCount = 0 Then 'exit sub
            GoTo end1
        End If

        'check for data
        strF = "ID_TBLSTUDIES = " & id_tblStudies
        rowsP = tblP.Select(strF)
        If rowsP.Length = 0 Then 'continue
        Else
            GoTo end1
        End If

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID As Int64
        Dim maxID1 As Int64

        maxID = GetMaxID("TBLTABLEPROPERTIES", 1, False) 'if maxid increment is 1, then getmaxid already does putmaxid
        maxID1 = maxID

        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If

        'strFMax = "charTable = 'TBLTABLEPROPERTIES'"
        'tblMax = tblMaxID
        'rowsMax = tblMax.Select(strFMax)
        'maxID = rowsMax(0).Item("nummaxid")
        'maxID1 = maxID

        If intRow = -1 Then
            intS = 0
            intE = dgv.RowCount - 1
        Else
            intS = intRow
            intE = intRow
        End If

        For Count1 = intS To intE
            idT = dgv("ID_TBLREPORTTABLE", Count1).Value
            idS = dgv("ID_TBLSTUDIES", Count1).Value
            idC = dgv("ID_TBLCONFIGREPORTTABLES", Count1).Value

            maxID = maxID + 1
            Dim row As DataRow = tblP.NewRow
            row.BeginEdit()
            row.Item("ID_TBLTABLEPROPERTIES") = maxID
            row.Item("ID_TBLREPORTTABLE") = idT
            row.Item("ID_TBLSTUDIES") = idS
            row.Item("ID_TBLCONFIGREPORTTABLES") = idC
            row.Item("BOOLBQLSHOWCONC") = -1
            row.Item("BOOLCSSHOWREJVALUES") = -1
            '20181220 LEE:
            'Change default for BOOLCSREPORTACCVALUES
            row.Item("BOOLCSREPORTACCVALUES") = 0 ' -1
            row.Item("BOOLQCREPORTACCVALUES") = -1
            row.Item("BOOLSTATSMEAN") = -1
            row.Item("BOOLSTATSSD") = -1
            row.Item("BOOLSTATSCV") = -1
            If boolAllowAcc(idC) Then
                row.Item("BOOLSTATSBIAS") = -1
            Else
                row.Item("BOOLSTATSBIAS") = 0
            End If
            row.Item("BOOLSTATSN") = -1
            row.Item("BOOLSTATSDIFF") = 0
            row.Item("BOOLSTATSDIFFCOL") = 0
            row.Item("BOOLSTATSREGR") = 0
            row.Item("BOOLSTATSNR") = -1
            row.Item("BOOLSTATSLETTER") = 0
            row.Item("BOOLTHEORETICAL") = 0

            row.Item("NUMISCRIT1") = 20 'System.DBNull.Value
            row.Item("NUMISCRIT1LEVEL") = System.DBNull.Value
            row.Item("NUMISCRIT2") = System.DBNull.Value
            row.Item("UPSIZE_TS") = "01-SEP-07"

            row.Item("BOOLINCLANOVA") = -1
            row.Item("BOOLINCLANOVASUMSTATS") = -1

            row.Item("BOOLBQLLEGEND") = 0
            row.Item("NUMSAMPLEG1") = 103
            row.Item("NUMSAMPLEG2") = 0
            row.Item("NUMSAMPLEG3") = 0
            row.Item("NUMSAMPLEG4") = 0
            row.Item("NUMSAMPLES1") = 103
            row.Item("NUMSAMPLES2") = 102
            row.Item("NUMSAMPLES3") = 106
            row.Item("NUMSAMPLES4") = 101
            row.Item("CHARSAMPLEGAD1") = "ASC"
            row.Item("CHARSAMPLEGAD2") = "ASC"
            row.Item("CHARSAMPLEGAD3") = "ASC"
            row.Item("CHARSAMPLEGAD4") = "ASC"
            row.Item("CHARSAMPLESAD1") = "ASC"
            row.Item("CHARSAMPLESAD2") = "ASC"
            row.Item("CHARSAMPLESAD3") = "ASC"
            row.Item("CHARSAMPLESAD4") = "ASC"
            row.Item("BOOLINCLUDEPSAE") = 0

            row.Item("BOOLRCCONC") = -1
            row.Item("BOOLRCPA") = 0
            row.Item("BOOLRCPARATIO") = 0
            row.Item("BOOLINCLUDEISTBL") = 0

            row.Item("BOOLNONELEG") = -1
            row.Item("BOOLPOSLEG") = 0
            row.Item("BOOLNEGLEG") = 0
            row.Item("BOOLCUSTOMLEG") = 0

            row.Item("CHARTITLELEG") = ""
            row.Item("CHARNUMLEG") = ""
            row.Item("CHARDENLEG") = ""

            row.Item("BOOLMEANACCURACY") = 0
            row.Item("BOOLINCLUDEDATE") = 0
            row.Item("BOOLDIFFERENCE") = -1
            row.Item("BOOLRECOVERY") = 0
            row.Item("BOOLINCLUDEWATSONLABELS") = 0
            row.Item("BOOLINTRARUNSUMSTATS") = 0

            row.Item("BOOLDOINDREC") = 0

            row.Item("BOOLREASSAYREASLETTERS") = 0

            row.Item("BOOLSTATSRE") = 0

            row.Item("CHARTIMEPERIOD") = System.DBNull.Value
            row.Item("CHARTIMEFRAME") = System.DBNull.Value
            row.Item("CHARPERIODTEMP") = System.DBNull.Value
            row.Item("INTNUMBEROFCYCLES") = System.DBNull.Value
            row.Item("BOOLCONVERTTIME") = -1
            row.Item("BOOLCONVERTTEMP") = -1
            row.Item("CHARISCONC") = System.DBNull.Value



            If boolAllowAcc(idC) Then
            Else
                row.Item("BOOLSTATSBIAS") = 0
                row.Item("BOOLTHEORETICAL") = 0
                row.Item("BOOLSTATSDIFF") = 0
                row.Item("BOOLSTATSDIFFCOL") = 0
                row.Item("BOOLSTATSRE") = 0
            End If

            row.EndEdit()
            tblP.Rows.Add(row)

        Next


        If boolGuWuOracle Then
            Try
                ta_tblTableProperties.Update(tblTableProperties)
            Catch ex As DBConcurrencyException
                'ds2005.TBLTABLEPROPERTIES.Merge('ds2005.TBLTABLEPROPERTIES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblTablePropertiesAcc.Update(tblTableProperties)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblTablePropertiesSQLServer.Update(tblTableProperties)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLTABLEPROPERTIES.Merge('ds2005Acc.TBLTABLEPROPERTIES, True)
            End Try
        End If


        If maxID = maxID1 Then
        Else

            Call PutMaxID("TBLTABLEPROPERTIES", maxID)

            'rowsMax(0).BeginEdit()
            'rowsMax(0).Item("nummaxid") = maxID + 1
            'rowsMax(0).EndEdit()

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


end1:

    End Sub

    Sub FillProperties()

        Dim dgv As DataGridView
        Dim tblP As System.Data.DataTable
        Dim rowsP() As DataRow
        Dim rowsIS() As DataRow
        Dim intRow As Short
        Dim idT As Int64
        Dim strF As String
        Dim var1, var2, var3
        Dim Count1 As Short
        Dim str1 As String
        Dim boolT As Boolean
        Dim idC As Int64

        tblP = tblTableProperties

        dgv = Me.dgvReportTables
        If dgv.RowCount = 0 Then 'exit sub
            Me.txtTitle.Text = ""
            GoTo end1
        End If
        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        'do txtTitle
        Me.txtTitle.Text = NZ(dgv("CHARHEADINGTEXT", intRow).Value, "")

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        idC = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

        'set public variable
        gidTR = idT
        gidCRT = idC

        rowsP = tblP.Select(strF, "", DataViewRowState.CurrentRows)
        If rowsP.Length = 0 Then
            Call InsertDefault(intRow)
        End If
        Erase rowsP 're-establish rowsp
        rowsP = tblP.Select(strF)

        boolT = boolHold
        boolHold = True 'set to true to disable check change event

        Dim intRows As Short
        intRows = rowsP.Length

        Dim boolC As Boolean
        Dim boolText As Boolean
        Dim boolNum As Boolean

        If intRows = 0 Then
            GoTo end1
        End If

        'do BQLShowConc
        var1 = rowsP(0).Item("BOOLBQLSHOWCONC")
        If var1 = -1 Then
            boolC = True
        Else
            boolC = False
        End If
        Me.rbShowBQL.Checked = boolC
        Me.rbDontShowBQL.Checked = Not (boolC)

        'do CalibrStd Show Rej Values
        var1 = rowsP(0).Item("BOOLCSSHOWREJVALUES")
        If var1 = -1 Then
            boolC = True
        Else
            boolC = False
        End If
        Me.rbShowRejectedValues.Checked = boolC
        Me.rbDontShowRejected.Checked = Not (boolC)

        'DO CalibrStd Show Accept Values
        var1 = rowsP(0).Item("BOOLCSREPORTACCVALUES")
        If var1 = -1 Then
            boolC = True
        Else
            boolC = False
        End If
        '20181220 LEE:
        'This is now chkRTC_CalStd_Acc
        'Me.rbRTC_CalStd_Acc.Checked = boolC
        'Me.rbRTC_CalStd_All.Checked = Not (boolC)
        Me.chkRTC_CalStd_Acc.Checked = boolC

        'do QC Show Accept Values
        var1 = rowsP(0).Item("BOOLQCREPORTACCVALUES")
        If var1 = -1 Then
            boolC = True
        Else
            boolC = False
        End If
        Me.rbRTC_QC_Acc.Checked = boolC
        Me.rbRTC_QC_All.Checked = Not (boolC)

        ''debugging
        'Dim col As DataColumn
        'For Each col In tblTableProperties.Columns
        '    ''''''''''console.writeline(col.ColumnName)
        'Next

        'do Stats Options
        For Count1 = 1 To 53
            boolText = False
            boolNum = False
            Select Case Count1
                Case 1
                    str1 = "BOOLSTATSMEAN"
                Case 2
                    str1 = "BOOLSTATSSD"
                Case 3
                    str1 = "BOOLSTATSCV"
                Case 4
                    str1 = "BOOLSTATSBIAS"
                Case 5
                    str1 = "BOOLSTATSN"
                Case 6
                    str1 = "BOOLSTATSDIFF"
                Case 7
                    str1 = "BOOLSTATSDIFFCOL"
                Case 8
                    str1 = "BOOLSTATSREGR"
                Case 9
                    str1 = "BOOLSTATSNR"
                    '- deprecated
                    '20181111 LEE:
                    'BOOLSTATSNR is used for stability type

                Case 10
                    str1 = "BOOLSTATSLETTER"
                Case 11
                    str1 = "BOOLTHEORETICAL"
                Case 12
                    str1 = "BOOLINCLANOVA"
                Case 13
                    str1 = "BOOLBQLLEGEND"
                Case 14
                    str1 = "BOOLINCLUDEPSAE"
                Case 15
                    str1 = "BOOLRCCONC"
                Case 16
                    str1 = "BOOLRCPA"
                Case 17
                    str1 = "BOOLRCPARATIO"
                Case 18
                    str1 = "BOOLINCLUDEISTBL"

                Case 19
                    str1 = "BOOLNONELEG"
                Case 20
                    str1 = "BOOLPOSLEG"
                Case 21
                    str1 = "BOOLNEGLEG"
                Case 22
                    str1 = "BOOLCUSTOMLEG"
                Case 23
                    str1 = "CHARTITLELEG"
                    boolText = True
                Case 24
                    str1 = "CHARNUMLEG"
                    boolText = True
                Case 25
                    str1 = "CHARDENLEG"
                    boolText = True
                Case 26
                    str1 = "BOOLMEANACCURACY"
                Case 27
                    str1 = "BOOLINCLUDEDATE"
                Case 28
                    str1 = "BOOLRECOVERY"
                Case 29
                    str1 = "BOOLDIFFERENCE"
                Case 30
                    str1 = "BOOLINCLUDEWATSONLABELS"
                Case 31
                    str1 = "BOOLINCLANOVASUMSTATS"
                Case 32
                    str1 = "BOOLSTATSRE"

                Case 33
                    str1 = "CHARTIMEPERIOD"
                    boolText = True
                Case 34
                    str1 = "CHARTIMEFRAME"
                    boolText = True
                Case 35
                    str1 = "CHARPERIODTEMP"
                    boolText = True
                Case 36
                    str1 = "INTNUMBEROFCYCLES"
                    boolNum = True
                Case 37
                    str1 = "BOOLCONVERTTIME"
                Case 38
                    str1 = "BOOLCONVERTTEMP"
                Case 39
                    str1 = "CHARISCONC"
                    boolText = True
                Case 40
                    str1 = "BOOLINTRARUNSUMSTATS"

                Case 41
                    str1 = "BOOLDOINDREC"

                Case 42
                    str1 = "BOOLREASSAYREASLETTERS"

                Case 43
                    str1 = "BOOLMFTABLE"
                Case 44
                    str1 = "BOOLINCLMFCOLS"
                Case 45
                    str1 = "BOOLINCLINTSTDNMF"
                Case 46
                    str1 = "BOOLCALCINTSTDNMF"
                Case 47
                    str1 = "CHARCARRYOVERLABEL"
                    boolText = True

                Case 48
                    str1 = "NUMPRECCRITLOTS"
                    boolNum = True


                Case 49
                    str1 = "BOOLREGRULOQ"

                Case 50
                    str1 = "INTQCLEVELGROUP"
                    boolNum = True

                Case 51
                    str1 = "BOOLISCOMBINELEVELS"

                Case 52
                    str1 = "BOOLCONCCOMMENTS"

                Case 53
                    str1 = "BOOLADHOCSTABCOMPCOLUMNS"

            End Select

            If boolText Then
                var1 = NZ(rowsP(0).Item(str1), "")
            ElseIf boolNum Then
                Select Case Count1
                    Case 48
                        var1 = NZ(rowsP(0).Item(str1), 0)
                    Case Else
                        var1 = NZ(rowsP(0).Item(str1), "")
                End Select
            Else
                var1 = NZ(rowsP(0).Item(str1), 0)
                If var1 = -1 Then
                    boolC = True
                Else
                    boolC = False
                End If
            End If

            'Select Case str1
            '    Case "CHARTITLELEG", "CHARNUMLEG", "CHARDENLEG"
            '        var1 = NZ(rowsP(0).Item(str1), "")
            '    Case Else
            '        var1 = NZ(rowsP(0).Item(str1), 0)
            '        If var1 = -1 Then
            '            boolC = True
            '        Else
            '            boolC = False
            '        End If
            'End Select

            Select Case Count1
                Case 1
                    Me.chkMean.Checked = boolC
                Case 2
                    Me.chkSD.Checked = boolC
                Case 3
                    Me.chkCV.Checked = boolC
                Case 4
                    Me.chkBias.Checked = boolC
                Case 5
                    Me.chkN.Checked = boolC
                Case 6
                    Me.chkDiff.Checked = boolC
                Case 7
                    Me.chkDiffCol.Checked = boolC
                Case 8
                    Me.chkRegr.Checked = boolC
                Case 9
                    Me.rbNR.Checked = boolC
                Case 10
                    Me.rbOutier.Checked = boolC
                Case 11
                    Me.chkTheoretical.Checked = boolC
                Case 12
                    Me.chkIncludeAnova.Checked = boolC
                Case 13
                    Me.chkBQLLEGEND.Checked = boolC
                Case 14
                    Me.chkIncludePSAE.Checked = boolC
                Case 15
                    Me.rbConc.Checked = boolC
                Case 16
                    Me.rbUsePeakArea.Checked = boolC
                Case 17
                    Me.rbUsePeakAreaRatio.Checked = boolC
                Case 18
                    Me.chkIncludeIS.Checked = boolC
                    Me.chkIncludeIS_Single.Checked = boolC
                Case 19
                    Me.chkNoneLeg.Checked = boolC
                Case 20
                    Me.rbPosLeg.Checked = boolC
                Case 21
                    Me.rbNegLeg.Checked = boolC
                Case 22
                    Me.chkCustomLeg.Checked = boolC
                Case 23
                    Me.CHARTITLELEG.Text = var1
                Case 24
                    Me.CHARNUMLEG.Text = var1
                Case 25
                    Me.CHARDENLEG.Text = var1
                Case 26
                    Me.rbMeanAcc.Checked = boolC
                Case 27
                    Me.chkIncludeDate.Checked = boolC
                Case 28
                    Me.rbRecovery.Checked = boolC
                Case 29
                    Me.rbDifference.Checked = boolC
                Case 30
                    Me.chkIncludeWatsonLabel.Checked = boolC
                Case 31
                    Me.chkIncludeAnovaSumStats.Checked = boolC
                Case 32
                    Me.chkRE.Checked = boolC

                Case 33
                    Me.CHARTIMEPERIOD.Text = var1
                Case 34
                    Me.CHARTIMEFRAME.Text = var1
                Case 35
                    Me.CHARPERIODTEMP.Text = var1
                Case 36
                    Me.INTNUMBEROFCYCLES.Text = var1
                Case 37
                    Me.chkCONVERTTIME.Checked = boolC
                Case 38
                    Me.chkCONVERTTEMP.Checked = boolC
                Case 39
                    Me.CHARISCONC.Text = var1
                Case 40
                    Me.chkIntraRunSumStats.Checked = boolC
                Case 41
                    Me.chkBOOLDOINDREC.Checked = boolC
                Case 42
                    Me.chkBOOLREASSAYREASLETTERS.Checked = boolC


                Case 43
                    Me.chkMFTable.Checked = boolC
                Case 44
                    Me.chkInclMFCols.Checked = boolC
                Case 45
                    Me.chkInclIntStdNMF.Checked = boolC
                Case 46
                    Me.chkCalcIntStdNMF.Checked = boolC

                    Try
                        If boolC Then
                            Me.rbOld.Checked = True
                            Me.rbNew.Checked = False
                        Else
                            Me.rbOld.Checked = False
                            Me.rbNew.Checked = True
                        End If
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    var1 = var1

                Case 47


                    '20181111 LEE:
                    'CHARCARRYOVERLABEL is also used for txtStabilityNotes

                    Select Case idC
                        Case 12, 18, 19, 21, 22, 23, 29, 31, 32 '20190220 LEE: Added 12=Dilution
                            If StrComp(var1, "Blank", CompareMethod.Text) = 0 Then
                                var1 = ""
                            End If
                            Me.txtStabilityNotes.Text = var1
                        Case Else
                            Me.CHARCARRYOVERLABEL.Text = var1
                    End Select

                    Me.txtStabilityNotes.Text = var1

                Case 48
                    Me.NUMPRECCRITLOTS.Text = var1

                Case 49
                    Me.chkBOOLREGRULOQ.Checked = boolC

                Case 50
                    var2 = NZ(var1, 0)
                    If var2 = 0 Then
                        Me.rbINTQCLEVELGROUPLevel.Checked = True
                        Me.rbINTQCLEVELGROUPNomConc.Checked = False
                        Me.rbINTQCLEVELGROUPQCLabel.Checked = False
                    ElseIf var2 = 1 Then
                        Me.rbINTQCLEVELGROUPLevel.Checked = False
                        Me.rbINTQCLEVELGROUPNomConc.Checked = True
                        Me.rbINTQCLEVELGROUPQCLabel.Checked = False
                    ElseIf var2 = 2 Then
                        Me.rbINTQCLEVELGROUPLevel.Checked = False
                        Me.rbINTQCLEVELGROUPNomConc.Checked = False
                        Me.rbINTQCLEVELGROUPQCLabel.Checked = True
                    End If

                Case 51
                    Me.chkBOOLISCOMBINELEVELS.Checked = boolC

                Case 52
                    Me.chkBOOLCONCCOMMENTS.Checked = boolC

                Case 53
                    Me.chkBOOLADHOCSTABCOMPCOLUMNS.Checked = boolC

            End Select
        Next

        '20190225 LEE:
        'rbUseISPeakArea logic
        If Me.rbUsePeakArea.Checked = False And Me.rbConc.Checked = False And Me.rbUsePeakAreaRatio.Checked = False Then
            Me.rbUseISPeakArea.Checked = True
        Else
            Me.rbUseISPeakArea.Checked = False
        End If

        'get charstabilityperiod
        'var1 = NZ(rowsP(0).Item("CHARSTABILITYPERIOD"), "")
        var1 = NZ(dgv("CHARSTABILITYPERIOD", intRow).Value, "")
        Me.CHARSTABILITYPERIOD.Text = var1
        'Me.CHARTIMEPERIOD.Text = ""
        'Me.CHARTIMEFRAME.Text = ""
        'Me.CHARPERIODTEMP.Text = ""

        'do groups and sorts
        Dim dtbl2 As System.Data.DataTable
        Dim str2 As String
        Dim boolU As Boolean = False

        rowsP(0).BeginEdit()
        var1 = rowsP(0).Item("NUMSAMPLEG1") 'this ID_TBLCONFIGHEADERLOOKUP
        If IsDBNull(var1) Then
            var1 = NZ(var1, 103)
            boolU = True
            rowsP(0).Item("NUMSAMPLEG1") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLEG1", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleG1.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleG1.Text = str1
        End Try

        var1 = rowsP(0).Item("NUMSAMPLEG2")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 0)
            boolU = True
            rowsP(0).Item("NUMSAMPLEG2") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLEG2", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleG2.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleG2.Text = str1
        End Try

        var1 = rowsP(0).Item("NUMSAMPLEG3")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 0)
            boolU = True
            rowsP(0).Item("NUMSAMPLEG3") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLEG3", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleG3.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleG3.Text = str1
        End Try

        var1 = rowsP(0).Item("NUMSAMPLEG4")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 0)
            boolU = True
            rowsP(0).Item("NUMSAMPLEG4") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLEG4", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleG4.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleG4.Text = str1
        End Try

        'Sort
        var1 = rowsP(0).Item("NUMSAMPLES1")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 103)
            boolU = True
            rowsP(0).Item("NUMSAMPLES1") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLES1", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleS1.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleS1.Text = str1
        End Try

        var1 = rowsP(0).Item("NUMSAMPLES2")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 102)
            boolU = True
            rowsP(0).Item("NUMSAMPLES2") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLES2", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleS2.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleS2.Text = str1
        End Try

        var1 = rowsP(0).Item("NUMSAMPLES3")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 106)
            boolU = True
            rowsP(0).Item("NUMSAMPLES3") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLES3", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleS3.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleS3.Text = str1
        End Try

        var1 = rowsP(0).Item("NUMSAMPLES4")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 101)
            boolU = True
            rowsP(0).Item("NUMSAMPLES4") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLES4", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleS4.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleS4.Text = str1
        End Try

        '****
        var1 = rowsP(0).Item("NUMSAMPLES5")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 101)
            boolU = True
            rowsP(0).Item("NUMSAMPLES5") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLES5", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleS5.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleS5.Text = str1
        End Try

        var1 = rowsP(0).Item("NUMSAMPLES6")
        If IsDBNull(var1) Then
            var1 = NZ(var1, 101)
            boolU = True
            rowsP(0).Item("NUMSAMPLES6") = var1
        End If
        str1 = RetrieveSG(var1)
        str2 = Replace("NUMSAMPLES6", "NUM", "CBX", 1, -1, CompareMethod.Text)
        Try
            Me.cbxSampleS6.Text = str1
        Catch ex As Exception
            str1 = "[None]"
            Me.cbxSampleS6.Text = str1

        End Try
        '****

        'Groups
        var1 = rowsP(0).Item("CHARSAMPLEGAD1")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLEGAD1") = var1
        End If
        str2 = Replace("CHARSAMPLEGAD1", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleGAD1.Text = var1

        var1 = rowsP(0).Item("CHARSAMPLEGAD2")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLEGAD2") = var1
        End If
        str2 = Replace("CHARSAMPLEGAD2", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleGAD2.Text = var1

        var1 = rowsP(0).Item("CHARSAMPLEGAD3")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLEGAD3") = var1
        End If
        str2 = Replace("CHARSAMPLEGAD3", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleGAD3.Text = var1

        var1 = rowsP(0).Item("CHARSAMPLEGAD4")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLEGAD4") = var1
        End If
        str2 = Replace("CHARSAMPLEGAD4", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleGAD4.Text = var1

        'Sorts
        var1 = rowsP(0).Item("CHARSAMPLESAD1")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLESAD1") = var1
        End If
        str2 = Replace("CHARSAMPLESAD1", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleSAD1.Text = var1

        var1 = rowsP(0).Item("CHARSAMPLESAD2")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLESAD2") = var1
        End If
        str2 = Replace("CHARSAMPLESAD2", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleSAD2.Text = var1

        var1 = rowsP(0).Item("CHARSAMPLESAD3")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLESAD3") = var1
        End If
        str2 = Replace("CHARSAMPLESAD3", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleSAD3.Text = var1

        var1 = rowsP(0).Item("CHARSAMPLESAD4")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLESAD4") = var1
        End If
        str2 = Replace("CHARSAMPLESAD4", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleSAD4.Text = var1

        '****

        var1 = rowsP(0).Item("CHARSAMPLESAD5")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLESAD5") = var1
        End If
        str2 = Replace("CHARSAMPLESAD5", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleSAD4.Text = var1

        var1 = rowsP(0).Item("CHARSAMPLESAD6")
        If IsDBNull(var1) Then
            var1 = NZ(var1, "ASC")
            boolU = True
            rowsP(0).Item("CHARSAMPLESAD6") = var1
        End If
        str2 = Replace("CHARSAMPLESAD6", "CHAR", "CBX", 1, -1, CompareMethod.Text)
        Me.cbxSampleSAD4.Text = var1

        '****

        'now set Stabilities
        Call FillStabilities()

        If boolU Then
            rowsP(0).EndEdit()
        End If

        Call DoMFTable()

        boolHold = boolT
end1:

        '20181112 LEE:
        Me.lblRemember.Visible = False

    End Sub

    Sub UpdateGroupSortData(ByVal ctr1 As Control)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim cbx As ComboBox
        Dim chk1 As System.Windows.Forms.CheckBox ' CheckBox
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim str1 As String
        Dim str2 As String
        Dim strC As String
        Dim strC1 As String
        Dim int1 As Int64
        Dim idL As Int64

        dtbl = tblTableProperties
        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)
        'idL = rows(0).Item("ID_TBLCONFIGHEADERLOOKUP")

        strC = ctr1.Name
        cbx = ctr1
        str1 = cbx.Text

        rows(0).BeginEdit()
        If InStr(1, ctr1.Name, "AD", CompareMethod.Text) > 0 Then
            strC1 = Replace(strC, "cbx", "CHAR", 1, -1, CompareMethod.Text)
            rows(0).Item(strC1) = str1
        Else
            strC1 = Replace(strC, "cbx", "NUM", 1, -1, CompareMethod.Text)
            int1 = PutSG(str1)
            rows(0).Item(strC1) = int1
        End If

        rows(0).EndEdit()


    End Sub

    Sub UpdateTextBoxes(ByVal tbx As TextBox)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim cbx As ComboBox
        Dim chk1 As System.Windows.Forms.CheckBox ' CheckBox
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim str1 As String
        Dim str2 As String
        Dim strC As String
        Dim strC1 As String
        Dim int1 As Int64
        Dim idL As Int64
        Dim var1
        Dim tName1 As String
        Dim tName2 As String
        Dim tName As String

        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT

        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)
        'idL = rows(0).Item("ID_TBLCONFIGHEADERLOOKUP")

        var1 = tbx.Text
        tName1 = Replace(tbx.Name, "QC", "", 1, -1, CompareMethod.Text)
        tName2 = Replace(tName1, "CAL", "", 1, -1, CompareMethod.Text)
        tName = tName2

        rows(0).BeginEdit()


        If IsDBNull(var1) Then
            rows(0).Item(tbx.Name) = DBNull.Value
        Else
            rows(0).Item(tName) = CSng(var1)
        End If

        rows(0).EndEdit()

    End Sub

    Sub DoLegendThings()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim rb1 As RadioButton
        Dim chk1 As System.Windows.Forms.CheckBox ' CheckBox
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim str1 As String
        Dim var1
        Dim boolPos As Boolean
        Dim boolCustom As Boolean
        Dim boolPeakArea As Boolean
        Dim strLegTitle As String
        Dim strLegNum As String
        Dim strLegDen As String
        Dim idRCT As Int64
        Dim boolMeanAcc As Boolean
        Dim boolNone As Boolean
        Dim boolRecovery As Boolean
        Dim boolDifference As Boolean
        Dim boolOnlyIS As Boolean

        dgv = Me.dgvReportTables
        'intRow = dgv.CurrentRow.Index


        If dgv.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
            dgv.CurrentCell = dgv.Item("CHARHEADINGTEXT", intRow)
            dgv.Rows(intRow).Selected = True
        Else
            intRow = dgv.CurrentRow.Index
        End If

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        idRCT = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)
        boolPos = Me.rbPosLeg.Checked
        boolOnlyIS = Me.chkCustomLeg.Checked
        boolPeakArea = Me.rbUsePeakArea.Checked
        boolMeanAcc = Me.rbMeanAcc.Checked
        boolNone = Me.chkNoneLeg.Checked
        boolRecovery = Me.rbRecovery.Checked
        boolDifference = Me.rbDifference.Checked

        'If boolCustom Then
        '    Me.CHARTITLELEG.Enabled = True
        '    Me.CHARNUMLEG.Enabled = True
        '    Me.CHARDENLEG.Enabled = True
        '    GoTo end1
        'ElseIf boolNone Then
        '    Me.CHARTITLELEG.Enabled = False
        '    Me.CHARNUMLEG.Enabled = False
        '    Me.CHARDENLEG.Enabled = False
        'Else
        '    Me.CHARTITLELEG.Enabled = False
        '    Me.CHARNUMLEG.Enabled = False
        '    Me.CHARDENLEG.Enabled = False
        'End If

        If boolNone Then
            'strLegTitle = ""
            'strLegNum = ""
            'strLegDen = ""
            'GoTo end1
        End If

        Select Case idRCT
            Case 13

                If boolPeakArea Then

                    If boolPos Then
                        strLegTitle = "Combined Recovery = "
                        strLegNum = "(Mean Peak Area of Extracted QC) * 100"
                        strLegDen = "(Mean Peak Area of Solution QC)"
                    Else
                        strLegTitle = "Combined Recovery = "
                        strLegNum = "(Mean Peak Area of Solution QC) * 100"
                        strLegDen = "(Mean Peak Area of Extracted QC)"
                    End If
                Else
                    If boolPos Then
                        strLegTitle = "Combined Recovery = "
                        strLegNum = "(Mean Peak Area Ratio of Extracted QC) * 100"
                        strLegDen = "(Mean Peak Area Ratio of Solution QC)"
                    Else
                        strLegTitle = "Combined Recovery = "
                        strLegNum = "(Mean Peak Area Ratio of Solution QC) * 100"
                        strLegDen = "(Mean Peak Area Ratio of Extracted QC)"
                    End If
                End If

            Case 14

                If boolPeakArea Then
                    If boolPos Then
                        strLegTitle = "True Recovery = "
                        strLegNum = "(Mean Peak Area of Extracted QC) * 100"
                        strLegDen = "(Mean Peak Area of Post Extraction Spiking Solution QC)"
                    Else
                        strLegTitle = "True Recovery = "
                        strLegNum = "(Mean Peak Area of Post Extraction Spiking Solution QC) * 100"
                        strLegDen = "(Mean Peak Area of Extracted QC)"
                    End If
                Else
                    If boolPos Then
                        strLegTitle = "True Recovery = "
                        strLegNum = "(Mean Peak Area Ratio of Extracted QC) * 100"
                        strLegDen = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                    Else
                        strLegTitle = "True Recovery = "
                        strLegNum = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC) * 100"
                        strLegDen = "(Mean Peak Area Ratio of Extracted QC)"
                    End If
                End If

            Case 15

                strLegTitle = DoLegendThingsMatrix(boolPos, "Title")
                strLegNum = DoLegendThingsMatrix(boolPos, "Num")
                strLegDen = DoLegendThingsMatrix(boolPos, "Den")

                'If BOOLMFTABLE Then
                '    If BOOLCALCINTSTDNMF Then
                '        If boolPos Then
                '            strLegTitle = "Internal Standard-Normalized Matrix Factor = "
                '            strLegNum = "(Analyte Matrix Factor)"
                '            strLegDen = "(Internal Standard Matrix Factor)"
                '        Else
                '            strLegTitle = "Matrix Factor = "
                '            strLegNum = "(Mean Peak Area Ratio of Solution QC) * 100"
                '            strLegDen = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                '        End If
                '    Else
                '        If boolPos Then
                '            strLegTitle = "Matrix Factor = "
                '            strLegNum = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                '            strLegDen = "(Mean Peak Area Ratio of Solution QC)"
                '        Else
                '            strLegTitle = "Matrix Factor = "
                '            strLegNum = "(Mean Peak Area Ratio of Solution QC)"
                '            strLegDen = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                '        End If
                '    End If
                'Else
                '    If boolPos Then
                '        strLegTitle = "Suppression/Enhancement = "
                '        strLegNum = "(Mean Peak Area of Post Extraction Spiking Solution QC)"
                '        strLegDen = "(Mean Peak Area of Solution QC)"
                '    Else
                '        strLegTitle = "Suppression/Enhancement = "
                '        strLegNum = "(Mean Peak Area of Solution QC)"
                '        strLegDen = "(Mean Peak Area of Post Extraction Spiking Solution QC)"
                '    End If
                'End If

                'If boolPeakArea Then
                '    If boolPos Then
                '        strLegTitle = "Suppression/Enhancement = "
                '        strLegNum = "(Mean Peak Area of Post Extraction Spiking Solution QC) * 100"
                '        strLegDen = "(Mean Peak Area of Solution QC)"
                '    Else
                '        strLegTitle = "Suppression/Enhancement = "
                '        strLegNum = "(Mean Peak Area of Solution QC) * 100"
                '        strLegDen = "(Mean Peak Area of Post Extraction Spiking Solution QC)"
                '    End If
                'Else
                '    If BOOLMFTABLE And BOOLCALCINTSTDNMF Then
                '        If boolPos Then
                '            strLegTitle = "Matrix Factor = "
                '            strLegNum = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC) * 100"
                '            strLegDen = "(Mean Peak Area Ratio of Solution QC)"
                '        Else
                '            strLegTitle = "Matrix Factor = "
                '            strLegNum = "(Mean Peak Area Ratio of Solution QC) * 100"
                '            strLegDen = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                '        End If
                '    Else
                '        If boolPos Then
                '            strLegTitle = "Matrix Factor = "
                '            strLegNum = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC) * 100"
                '            strLegDen = "(Mean Peak Area Ratio of Solution QC)"
                '        Else
                '            strLegTitle = "Matrix Factor = "
                '            strLegNum = "(Mean Peak Area Ratio of Solution QC) * 100"
                '            strLegDen = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                '        End If
                '    End If

                'End If

            Case 17

                strLegTitle = ""
                strLegNum = ""
                strLegDen = ""

            Case 22, 23, 29, 31, 32

                Call DoCalcsLegend()


                'Case 32

                '    strLegTitle = ""
                '    strLegNum = ""
                '    strLegDen = ""

        End Select

end2:

        Select Case idRCT
            Case 22, 23, 29, 31, 32
            Case Else
                If Len(Me.CHARTITLELEG.Text) = 0 Then
                    Me.CHARTITLELEG.Text = strLegTitle
                End If
                If Len(Me.CHARNUMLEG.Text) = 0 Then
                    Me.CHARNUMLEG.Text = strLegNum
                End If
                If Len(Me.CHARDENLEG.Text) = 0 Then
                    Me.CHARDENLEG.Text = strLegDen
                End If

        End Select

end1:

    End Sub

    Function DoLegendThingsMatrix(boolPos As Boolean, strType As String) As String

        Dim boolCalcIntStdNMF As Boolean = False

        Dim idC As Int32 = idCGet()

        Select Case idC
            Case 13, 14, 15
                Select Case strType
                    Case "Title"
                        If BOOLMFTABLE Then
                            If boolCalcIntStdNMF Then
                                If boolPos Then
                                    DoLegendThingsMatrix = "Internal Standard-Normalized Matrix Factor = "
                                Else
                                    DoLegendThingsMatrix = "Internal Standard-Normalized Matrix Factor = "
                                End If
                            Else
                                If boolPos Then
                                    DoLegendThingsMatrix = "Internal Standard-Normalized Matrix Factor = "
                                Else
                                    DoLegendThingsMatrix = "Internal Standard-Normalized Matrix Factor = "
                                End If
                            End If
                        Else
                            If boolPos Then
                                DoLegendThingsMatrix = "Suppression/Enhancement = "
                            Else
                                DoLegendThingsMatrix = "Suppression/Enhancement = "
                            End If
                        End If
                    Case "Num"
                        If BOOLMFTABLE Then
                            If boolCalcIntStdNMF Then
                                If boolPos Then
                                    DoLegendThingsMatrix = "(Analyte Matrix Factor)"
                                Else
                                    DoLegendThingsMatrix = "(Internal Standard Matrix Factor)"
                                End If
                            Else
                                If boolPos Then
                                    DoLegendThingsMatrix = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                                Else
                                    DoLegendThingsMatrix = "(Mean Peak Area Ratio of Solution QC)"
                                End If
                            End If
                        Else
                            If boolPos Then
                                DoLegendThingsMatrix = "(Mean Peak Area of Post Extraction Spiking Solution QC)"
                            Else
                                DoLegendThingsMatrix = "(Mean Peak Area of Solution QC)"
                            End If
                        End If
                    Case "Den"
                        If BOOLMFTABLE Then
                            If boolCalcIntStdNMF Then
                                If boolPos Then
                                    DoLegendThingsMatrix = "(Internal Standard Matrix Factor)"
                                Else
                                    DoLegendThingsMatrix = "(Analyte Matrix Factor)"
                                End If
                            Else
                                If boolPos Then
                                    DoLegendThingsMatrix = "(Mean Peak Area Ratio of Solution QC)"
                                Else
                                    DoLegendThingsMatrix = "(Mean Peak Area Ratio of Post Extraction Spiking Solution QC)"
                                End If
                            End If
                        Else
                            If boolPos Then
                                DoLegendThingsMatrix = "(Mean Peak Area of Solution QC)"
                            Else
                                DoLegendThingsMatrix = "(Mean Peak Area of Post Extraction Spiking Solution QC)"
                            End If
                        End If
                End Select
        End Select


    End Function

    Sub DoCalcsLegend()

        Dim boolMeanAcc As Boolean
        Dim boolNone As Boolean
        Dim boolRecovery As Boolean
        Dim boolDifference As Boolean
        Dim strLegTitle As String
        Dim strLegNum As String
        Dim strLegDen As String
        Dim boolPos As Boolean

        Dim strNum As String
        Dim strDenom As String
        Dim boolOld As Boolean

        boolPos = Me.rbPosLeg.Checked
        boolMeanAcc = Me.rbMeanAcc.Checked
        boolNone = Me.chkNoneLeg.Checked
        boolRecovery = Me.rbRecovery.Checked
        boolDifference = Me.rbDifference.Checked
        boolOld = Me.rbOld.Checked

        Dim strMA As String
        Dim strRec As String
        Dim strDiff As String
        Dim strPosLeg As String
        Dim strNegLeg As String

        strPosLeg = "(Old - New)"
        strNegLeg = "(New - Old)"

        If boolMeanAcc Then
            If boolPos Then
                strLegTitle = "Mean Accuracy = "
                strLegNum = "(Mean Old - Mean New) * 100"
                strLegDen = "(Mean New + Mean Old)/2"
            Else
                strLegTitle = "Mean Accuracy = "
                strLegNum = "(Mean New - Mean Old) * 100"
                strLegDen = "(Mean New + Mean Old)/2"
            End If
        ElseIf boolRecovery Then
            strPosLeg = "Old"
            strNegLeg = "New"
            If boolOld Then
                If boolPos Then
                    strLegTitle = "%Recovery = "
                    strLegNum = "(Mean Old) * 100"
                    strLegDen = "(Mean Old)"
                Else
                    strLegTitle = "%Recovery = "
                    strLegNum = "(Mean New) * 100"
                    strLegDen = "(Mean Old)"
                End If
            Else
                If boolPos Then
                    strLegTitle = "%Recovery = "
                    strLegNum = "(Mean Old) * 100"
                    strLegDen = "(Mean New)"
                Else
                    strLegTitle = "%Recovery = "
                    strLegNum = "(Mean New) * 100"
                    strLegDen = "(Mean New)"
                End If
            End If
         
        ElseIf boolDifference Then
            strLegTitle = "%Difference = "
            If boolOld Then
                If boolPos Then
                    strLegNum = "(Mean Old - Mean New) * 100"
                    strLegDen = "(Mean Old)"
                    strDiff = "(Old - New)/Old * 100"
                Else
                    strLegNum = "(Mean New - Mean Old) * 100"
                    strLegDen = "(Mean Old)"
                    strDiff = "(New - Old)/Old * 100"
                End If
            Else
                If boolPos Then
                    strLegNum = "(Mean Old - Mean New) * 100"
                    strLegDen = "(Mean New)"
                    strDiff = "(Old - New)/New * 100"
                Else
                    strLegNum = "(Mean New - Mean Old) * 100"
                    strLegDen = "(Mean New)"
                    strDiff = "(New - Old)/New * 100"
                End If
            End If


        Else
            If boolPos Then
                strLegTitle = "%Difference = "
                strLegNum = "(Mean Old - Mean New) * 100"
                strLegDen = "(Mean Old)"
            Else
                strLegTitle = "%Difference = "
                strLegNum = "(Mean New - Mean Old) * 100"
                strLegDen = "(Mean Old)"
            End If
        End If


        'now do lblDiff
        If boolOld Then
            If boolPos Then
                strMA = "(Old - New)/" & ChrW(10) & "((Old + New)/2) * 100"
                strRec = "Old/Old * 100"
                strDiff = "(Old - New)/Old * 100"
            Else
                strMA = "(New - Old)/" & ChrW(10) & "((Old + New)/2) * 100"
                strRec = "New/Old * 100"
                strDiff = "(New - Old)/Old * 100"
            End If
        Else
            If boolPos Then
                strMA = "(Old - New)/" & ChrW(10) & "((Old + New)/2) * 100"
                strRec = "Old/New * 100"
                strDiff = "(Old - New)/New * 100"
            Else
                strMA = "(New - Old)/" & ChrW(10) & "((Old + New)/2) * 100"
                strRec = "New/New * 100"
                strDiff = "(New - Old)/New * 100"
            End If
        End If

        Me.lblPercDiff.Text = strDiff
        Me.lblRecovery.Text = strRec
        Me.lblMeanAccuracy.Text = strMA

        Me.rbPosLeg.Text = strPosLeg
        Me.rbNegLeg.Text = strNegLeg

        '20181010 LEE:
        'Modify this logic

        If Len(Me.CHARTITLELEG.Text) = 0 Then
            Me.CHARTITLELEG.Text = strLegTitle
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
        End If

        If Len(Me.CHARNUMLEG.Text) = 0 Then
            Me.CHARNUMLEG.Text = strLegNum
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
        End If

        If Len(Me.CHARDENLEG.Text) = 0 Then
            Me.CHARDENLEG.Text = strLegDen
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")
        End If

        'Me.CHARTITLELEG.Text = strLegTitle
        'Me.CHARNUMLEG.Text = strLegNum
        'Me.CHARDENLEG.Text = strLegDen

        'Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
        'Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
        'Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")

    End Sub


    Sub UpdateTBData(ByVal tb As TextBox, ByVal strC As String)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim rb1 As RadioButton
        Dim chk1 As System.Windows.Forms.CheckBox ' CheckBox
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim str1 As String
        Dim var1

        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)

        var1 = tb.Text
        str1 = NZ(var1, "")

        rows(0).BeginEdit()
        Select Case strC
            Case "INTNUMBEROFCYCLES"
                If Len(str1) = 0 Then
                    rows(0).Item(strC) = System.DBNull.Value
                Else
                    rows(0).Item(strC) = str1
                End If

            Case Else
                rows(0).Item(strC) = str1
        End Select

        rows(0).EndEdit()

        Call SetTablePropertiesBool(gidTR, gidCRT)

    End Sub

    Sub UpdateChkData(ByVal boolRB As Boolean, ByVal ctr1 As Control, ByVal strC As String)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim rb1 As RadioButton
        Dim chk1 As System.Windows.Forms.CheckBox ' CheckBox
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim intRB As Short
        Dim str1 As String

        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)

        If Me.panNomDenom.Visible And Me.gbDenom.Visible And StrComp(strC, "BOOLCALCINTSTDNMF", CompareMethod.Text) = 0 Then
            If Me.rbOld.Checked Then
                Me.chkCalcIntStdNMF.Checked = True
                BOOLCALCINTSTDNMF = True
            Else
                Me.chkCalcIntStdNMF.Checked = False
                BOOLCALCINTSTDNMF = False
            End If

            boolC = BOOLCALCINTSTDNMF
        Else
            If boolRB Then
                rb1 = ctr1
                boolC = rb1.Checked
            Else
                chk1 = ctr1
                boolC = chk1.Checked
            End If
        End If

        rows(0).BeginEdit()
        '20181111 LEE: Account for gbStabilities
        Dim boolDo As Boolean = True

        If boolRB Then
            Select Case gidCRT
                Case 12, 18, 19, 21, 22, 23, 29, 31, 32

                    '20181111 LEE
                    'Stored in BOOLSTATSNR

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
                    '12   rbDilution    Dilution Samples

                    str1 = ""
                    Select Case UCase(rb1.Name)
                        Case "RBNA"
                            intRB = 1
                            str1 = ""
                        Case "RBPROCESS"
                            intRB = 2
                            str1 = "Extract (or processed sample) stability: The sponsor should assess the stability of processed samples, including the residence time in the autosampler against freshly prepared calibrators."
                        Case "RBBENCHTOP"
                            intRB = 3
                            str1 = "Bench-top stability: The sponsor should determine the stability of samples under the laboratory handling conditions that are expected for the study samples (e.g., the stability of samples maintained at room temperature or stored in an ice bucket)."
                        Case "RBFT"
                            intRB = 4
                            str1 = "Freeze-thaw stability: The sponsor should assess the stability of the sample after a minimum of three freeze-thaw cycles. QC samples should be thawed and analyzed according to the same procedures as the study samples." & ChrW(10) & "QC samples should be frozen for at least 12 hours between cycles. Freeze-thaw stability QCs should be compared to freshly prepared calibration curves and QCs."
                        Case "RBLT"
                            intRB = 5
                            str1 = "Long-term stability: The sponsor should determine the long-term stability of the sample over a period of time" & ChrW(10) & "equal to or exceeding the time between the date of first sample collection and the date of last sample analysis. The storage temperatures studied should be the same as those used to store study samples. Long-term stability QCs should be compared to freshly prepared calibration curves and QCs. Determination of stability at minus 20ºC would cover stability at colder temperatures."
                        Case "RBREINJECTION"
                            intRB = 6
                            str1 = "Reinjection Stability"
                        Case "RBBLOOD"
                            intRB = 7
                            str1 = "Whole Blood Stability"
                        Case "RBSTOCKSOLUTION"
                            intRB = 8
                            str1 = "Stock solution stability: Stock solutions should not be made from reference materials that are about to expire unless the purity of the analyte in the stock solutions is re-established." & ChrW(10) & "When the stock solution exists in a different state (e.g., solution versus solid) or in a different buffer composition (which is generally the case for macromolecules) from the certified reference standard, the sponsor should generate stability data on stock solutions to justify the duration of stock solution storage stability."
                        Case "RBSPIKING"
                            intRB = 9
                            str1 = "Stock solution stability: Stock solutions should not be made from reference materials that are about to expire unless the purity of the analyte in the stock solutions is re-established." & ChrW(10) & "When the stock solution exists in a different state (e.g., solution versus solid) or in a different buffer composition (which is generally the case for macromolecules) from the certified reference standard, the sponsor should generate stability data on stock solutions to justify the duration of stock solution storage stability."
                            '20190109 LEE:
                        Case "RBAUTOSAMPLER"
                            intRB = 10
                            str1 = "Autosampler stability: The sponsor should demonstrate the stability of extracts in the autosampler only if the autosampler storage conditions are different or not covered by extract (processed sample) stability."
                        Case "RBBATCHREINJECTION"
                            intRB = 11
                            str1 = "Batch Reinjection Stability"
                        Case "RBDILUTION"
                            intRB = 12
                            str1 = "Dilution Samples"
                        Case Else
                            boolDo = False
                    End Select

                    If boolDo Then
                        'only record if checked
                        If boolC Then
                            rows(0).Item(strC) = intRB
                        End If

                        Me.txtRef.Text = str1

                    Else
                        If boolC Then
                            rows(0).Item(strC) = -1
                        Else
                            rows(0).Item(strC) = 0
                        End If
                    End If


                Case Else

                    If boolC Then
                        rows(0).Item(strC) = -1
                    Else
                        rows(0).Item(strC) = 0
                    End If
            End Select
        Else
            If boolC Then
                rows(0).Item(strC) = -1
            Else
                rows(0).Item(strC) = 0
            End If
        End If
        
        rows(0).EndEdit()

        Call SetTablePropertiesBool(gidTR, gidCRT)

        'do some visibility things
        Call StabilityViews()


    End Sub

    Sub StabilityViews()

        If boolFormLoad Then
            Exit Sub
        End If


        Dim intRow As Short
        Dim dgv As DataGridView
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim int1 As Short
        Dim var1

        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)
        If rows.Length = 0 Then
        Else
            Try
                int1 = rows(0).Item("BOOLSTATSNR")
                If int1 = 12 Then
                    Me.gbAdditional.Visible = False
                    Me.panFDARef.Visible = False
                End If
            Catch ex As Exception
                var1 = var1
            End Try

        End If

    End Sub

    Sub UpdateFDARef()

        Dim str1 As String
        Dim str2 As String
        Dim Count1 As Short
        Dim intRB As Short
        Dim boolDo As Boolean
        Dim rb As RadioButton
        Dim boolShowFDA As Boolean = True

        Me.txtRef.Text = ""

        For Count1 = 1 To 12

            Select Case Count1
                Case 1
                    intRB = 1
                    str1 = ""
                    str2 = "RBNA"
                Case 2
                    intRB = 2
                    str1 = "Extract (or processed sample) stability: The sponsor should assess the stability of processed samples, including the residence time in the autosampler against freshly prepared calibrators."
                    str2 = "RBPROCESS"
                Case 3
                    intRB = 3
                    str1 = "Bench-top stability: The sponsor should determine the stability of samples under the laboratory handling conditions that are expected for the study samples (e.g., the stability of samples maintained at room temperature or stored in an ice bucket)."
                    str2 = "RBBENCHTOP"
                Case 4
                    intRB = 4
                    str1 = "Freeze-thaw stability: The sponsor should assess the stability of the sample after a minimum of three freeze-thaw cycles. QC samples should be thawed and analyzed according to the same procedures as the study samples." & ChrW(10) & "QC samples should be frozen for at least 12 hours between cycles. Freeze-thaw stability QCs should be compared to freshly prepared calibration curves and QCs."
                    str2 = "RBFT"
                Case 5
                    intRB = 5
                    str1 = "Long-term stability: The sponsor should determine the long-term stability of the sample over a period of time" & ChrW(10) & "equal to or exceeding the time between the date of first sample collection and the date of last sample analysis. The storage temperatures studied should be the same as those used to store study samples. Long-term stability QCs should be compared to freshly prepared calibration curves and QCs. Determination of stability at minus 20ºC would cover stability at colder temperatures."
                    str2 = "RBLT"
                Case 6
                    intRB = 6
                    str1 = "Reinjection Stability"
                    str2 = "RBREINJECTION"
                Case 7
                    intRB = 7
                    str1 = "Whole Blood Stability"
                    str2 = "RBBLOOD"
                Case 8
                    intRB = 8
                    str1 = "Stock solution stability: Stock solutions should not be made from reference materials that are about to expire unless the purity of the analyte in the stock solutions is re-established." & ChrW(10) & "When the stock solution exists in a different state (e.g., solution versus solid) or in a different buffer composition (which is generally the case for macromolecules) from the certified reference standard, the sponsor should generate stability data on stock solutions to justify the duration of stock solution storage stability."
                    str2 = "RBSTOCKSOLUTION"
                Case 9
                    intRB = 9
                    str1 = "Stock solution stability: Stock solutions should not be made from reference materials that are about to expire unless the purity of the analyte in the stock solutions is re-established." & ChrW(10) & "When the stock solution exists in a different state (e.g., solution versus solid) or in a different buffer composition (which is generally the case for macromolecules) from the certified reference standard, the sponsor should generate stability data on stock solutions to justify the duration of stock solution storage stability."
                    str2 = "RBSPIKING"
                    '20190109 LEE:
                Case 10
                    intRB = 10
                    str1 = "Autosampler stability: The sponsor should demonstrate the stability of extracts in the autosampler only if the autosampler storage conditions are different or not covered by extract (processed sample) stability."
                    str2 = "RBAUTOSAMPLER"
                Case 11
                    intRB = 11
                    str1 = "Batch Reinjection Stability"
                    str2 = "RBBATCHREINJECTION"
                Case 12
                    intRB = 12
                    str1 = "Dilution Samples"
                    str2 = "RBDILUTION"
                Case Else
                    boolDo = False
            End Select

            Try
                rb = Me.gbStabilityType.Controls(str2)
                If rb.Checked Then
                    Me.txtRef.Text = str1
                    Exit For
                End If
            Catch ex As Exception

            End Try

        Next Count1

        Me.txtRef.Text = str1

        'Me.panFDARef.Visible = Me.gbStabilityType.Visible

        '20190220 LEE:
        Dim idC As Int32 = idCGet()
        Select Case idC
            Case 18, 19, 21, 22, 23, 29, 31, 32 'do not show for 12=Dilution
                If idC = 31 Then 'check in case assay is dilution
                    Call StabilityViews()
                Else
                    Me.panFDARef.Visible = True
                End If

            Case Else
                Me.panFDARef.Visible = False
        End Select

    End Sub

    Sub UpdateNumeric(ByVal ctr1 As Control, ByVal strC As String)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim rb1 As RadioButton
        Dim chk1 As System.Windows.Forms.CheckBox ' CheckBox
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim num1 As Decimal
        Dim strName As String

        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)

        strName = ctr1.Name
        If InStr(1, strName, "INTQCLEVELGROUP", CompareMethod.Text) > 0 Then
            If InStr(1, strName, "GROUPLevel", CompareMethod.Text) > 0 Then
                num1 = 0
            ElseIf InStr(1, strName, "NomConc", CompareMethod.Text) > 0 Then
                num1 = 1
            ElseIf InStr(1, strName, "QCLabel", CompareMethod.Text) > 0 Then
                num1 = 2
            End If
        Else
            num1 = ctr1.Text
        End If


        rows(0).BeginEdit()
        rows(0).Item(strC) = num1
        rows(0).EndEdit()

        Call SetTablePropertiesBool(gidTR, gidCRT)


    End Sub

    Sub UpdatecharStabilityPeriod()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolC As Boolean
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64
        Dim str1 As String

        str1 = Me.CHARSTABILITYPERIOD.Text

        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index

        idT = dgv("ID_TBLREPORTTABLE", intRow).Value
        dtbl = tblReportTable
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT
        rows = dtbl.Select(strF)

        rows(0).BeginEdit()
        rows(0).Item("CHARSTABILITYPERIOD") = str1
        rows(0).EndEdit()

        'must do this too
        dgv("CHARSTABILITYPERIOD", intRow).Value = str1

        Call SetTablePropertiesBool(gidTR, gidCRT)


    End Sub

    Sub DoRTCancel()


        'reset dgv from arrbackup
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim dgv As DataGridView = Me.dgvReportTables
        Dim dv As System.Data.DataView = dgv.DataSource
        Dim dtbl As DataTable = tblReportTables
        Dim id1 As Int64
        Dim id2 As Int64
        Dim int1 As Int16
        Dim int2 As Int16

        For Count1 = 0 To dtbl.Rows.Count - 1
            dtbl.Rows(Count2).BeginEdit()
            For Count2 = 0 To dtbl.Columns.Count - 1
                dtbl.Rows(Count1).Item(Count2) = arrBU(Count2 + 1, Count1 + 1)
            Next
            dtbl.Rows(Count2).EndEdit()
        Next

        ''For Count1 = 0 To Me.dgvReportTables.Rows.Count - 1
        'For Count1 = 0 To dv.Count - 1
        '    dv(Count1).BeginEdit()
        '    For Count2 = 0 To Me.dgvReportTables.Columns.Count - 1
        '        dv(Count1).Item(Count2) = arrBU(Count1, Count2)
        '    Next
        '    dv(Count1).EndEdit()
        'Next

    End Sub

    Sub DoPropCancel()

        tblTableProperties.RejectChanges()

        Call DoRTCancel()

        Call FillProperties()

    End Sub

    Sub DoASPCancel()

        tblAutoAssignSamples.RejectChanges()
        Me.dgvASP.Refresh()
        Call FilterSAS()

    End Sub

    Sub DoSASCancel()

        Me.dgvSAS.CancelEdit()
        tblSAS.RejectChanges()
        'Call FilterSAS()

    End Sub

    Private Sub rbShowBQL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbShowBQL.CheckedChanged
        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbShowBQL, "BOOLBQLSHOWCONC")
        End If
    End Sub

    Private Sub rbShowRejectedValues_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbShowRejectedValues.CheckedChanged

        '20181108 LEE: gbxSuper has been depricated
        'If Me.rbShowRejectedValues.Checked Then
        '    Me.gbxSuper.Enabled = True
        'Else
        '    Me.gbxSuper.Enabled = False
        'End If

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbShowRejectedValues, "BOOLCSSHOWREJVALUES")
        End If

    End Sub

    Private Sub rbRTC_CalStd_Acc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRTC_CalStd_Acc.CheckedChanged

        '20181220 LEE:
        'Don't do this anymore. Radiobutton is deprecated
        'BOOLCSREPORTACCVALUES will be used for Do Calculations True/False
        'If boolHold Then
        'Else
        '    Call UpdateChkData(True, Me.rbRTC_CalStd_Acc, "BOOLCSREPORTACCVALUES")
        'End If

    End Sub

    Private Sub rbRTC_QC_Acc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRTC_QC_Acc.CheckedChanged
        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbRTC_QC_Acc, "BOOLQCREPORTACCVALUES")
        End If

    End Sub

    Private Sub chkMean_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMean.CheckedChanged
        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkMean, "BOOLSTATSMEAN")
            Call CheckStats(Me.chkMean)
        End If
    End Sub

    Private Sub chkSD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSD.CheckedChanged
        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkSD, "BOOLSTATSSD")
            Call CheckStats(Me.chkSD)
        End If
    End Sub

    Private Sub chkCV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCV.CheckedChanged
        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkCV, "BOOLSTATSCV")
            Call CheckStats(Me.chkCV)
        End If

    End Sub

    Private Sub chkBias_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBias.CheckedChanged

        Dim bool As Boolean
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id As Short
        Dim strM As String

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        bool = Me.chkBias.Checked


        If bool Then
            dgv = Me.dgvReportTables
            intRow = dgv.CurrentRow.Index
            id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

            Dim bool1 As Boolean
            If boolAllowAcc(id) Then
            Else
                strM = "This table does not support the %Bias parameter."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkBias.Checked = False
            End If
        End If

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkBias, "BOOLSTATSBIAS")
            Call CheckStats(Me.chkBias)
        End If


    End Sub

    Private Sub chkN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkN.CheckedChanged
        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkN, "BOOLSTATSN")
            Call CheckStats(Me.chkN)
        End If

    End Sub

    Sub CheckStats(ByVal chk1 As System.Windows.Forms.CheckBox)

        Dim boolM As Boolean
        Dim boolSD As Boolean
        Dim boolCV As Boolean
        Dim boolBias As Boolean
        Dim boolRegr As Boolean
        Dim boolDiff As Boolean
        Dim boolDiffCol As Boolean
        Dim boolTheoretical As Boolean
        Dim boolRE As Boolean
        Dim strM As String

        boolM = Me.chkMean.Checked
        boolSD = Me.chkSD.Checked
        boolCV = Me.chkCV.Checked
        boolBias = Me.chkBias.Checked
        boolRegr = Me.chkRegr.Checked
        boolDiff = Me.chkDiff.Checked
        boolDiffCol = Me.chkDiffCol.Checked
        boolTheoretical = Me.chkTheoretical.Checked
        boolRE = Me.chkRE.Checked

        'if chkSD = true, then chkMean must be checked
        If boolSD Then
            If boolM Then 'ignore
            Else
                strM = "Since SD is selected, Mean must also be selected."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkMean.Checked = True
            End If
        End If


        'if chkCV = true, then chkMean and chkSD must be checked
        If boolCV Then
            If boolM Then 'ignore
            Else
                If StrComp(chk1.Name, "chkMean", CompareMethod.Text) = 0 Then
                    strM = "Since " & ReturnPrecLabel() & " is selected, Mean must also be selected."
                    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                End If
                Me.chkMean.Checked = True
            End If
            If boolSD Then 'ignore
            Else
                strM = "Since " & ReturnPrecLabel() & " is selected, SD must also be selected."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkSD.Checked = True
            End If
        End If

        'if chkBias = true or chkTheor, then chkMean must be checked
        '20160905 LEE: Disable the Accuracy settings to allow users to show individual Accuracy column without having to show stats section
        GoTo end1

        If boolBias Then
            If boolM Then 'ignore
            Else
                strM = "Since %Bias is selected, Mean must also be selected."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkMean.Checked = True
            End If
        End If

        If boolTheoretical Then
            If boolM Then 'ignore
            Else
                strM = "Since %Theoretical is selected, Mean must also be selected."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkMean.Checked = True
            End If
        End If

        If boolDiff Then
            If boolM Then 'ignore
            Else
                strM = "Since %Diff is selected, Mean must also be selected."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkMean.Checked = True
            End If
        End If

        If boolRE Then
            If boolM Then 'ignore
            Else
                strM = "Since %RE is selected, Mean must also be selected."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkMean.Checked = True
            End If
        End If

end1:

    End Sub

    Private Sub cmdIncSamples_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIncSamples.Click

        Dim frm As New frmIncSmplCrit
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim var1, var2, var3
        Dim strF As String
        Dim dtbl As System.Data.DataTable
        Dim row() As DataRow
        Dim idR As Int64


        dgv = Me.dgvReportTables
        intRow = dgv.CurrentRow.Index
        idR = dgv("ID_TBLREPORTTABLE", intRow).Value
        strF = "ID_TBLREPORTTABLE = " & idR & " AND ID_TBLSTUDIES = " & id_tblStudies
        dtbl = tblTableProperties
        row = dtbl.Select(strF)

        var1 = NZ(row(0).Item("NUMISCRIT1"), "")
        frm.NUMISCRIT1.Text = CStr(var1)

        var2 = NZ(row(0).Item("NUMISCRIT1LEVEL"), "")
        frm.NUMISCRIT1LEVEL.Text = CStr(var2)

        var3 = NZ(row(0).Item("NUMISCRIT2"), "")
        frm.NUMISCRIT2.Text = CStr(var3)

        If Len(var2) = 0 Or Len(var3) = 0 Then
            frm.rb1.Checked = True
            frm.rb2.Checked = False
        Else
            frm.rb1.Checked = False
            frm.rb2.Checked = True
        End If

        frm.ShowDialog()

        If frm.boolCancel Then 'ignore changes
        Else
            row(0).BeginEdit()

            var1 = frm.NUMISCRIT1.Text
            row(0).Item("NUMISCRIT1") = var1

            var1 = NZ(frm.NUMISCRIT1LEVEL.Text, System.DBNull.Value)
            row(0).Item("NUMISCRIT1LEVEL") = var1

            var1 = NZ(frm.NUMISCRIT2.Text, System.DBNull.Value)
            row(0).Item("NUMISCRIT2") = var1

            row(0).EndEdit()

        End If

        frm.Dispose()

    End Sub

    Private Sub cmdLegend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("Under construction")
        Exit Sub


    End Sub


    Sub BuildPeriodTemp(ByVal boolFromInsert As Boolean)

        Dim var1, var2, var3, var1a
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strT As String
        Dim intS As Short
        Dim intE As Short
        Dim strF As String
        Dim strM As String
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim dv As System.Data.DataView
        Dim boolC As Boolean
        Dim varC
        Dim strC As String
        Dim strTP As String
        Dim int1 As Short
        Dim strFC As String
        Dim strFTP As String

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled = False Then
            Exit Sub
        End If

        Dim boolNoC As Boolean = False
        Dim boolNoTP As Boolean = False

        strM = ""
        'boolC = CheckStabilityValidation()
        boolC = True
        'If boolC Then
        'Else
        '    GoTo end1
        'End If

        dgv = Me.dgvReportTables
        dv = dgv.DataSource
        intRow = dgv.CurrentRow.Index

        'check text box validations

        'If Me.panTP.Visible Then

        varC = Me.INTNUMBEROFCYCLES.Text
        var1 = Me.CHARTIMEPERIOD.Text
        var2 = Me.CHARTIMEFRAME.Text
        var3 = Me.CHARPERIODTEMP.Text

        If Len(varC) = 0 Then
            strC = ""
        Else
            str1 = varC & " Cycles"
            str2 = Replace(str1, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
            strC = str2
        End If

        'get number for var1
        If Len(var1) = 0 Then
            var1a = ""
        Else
            If IsNumeric(var1) Then
                If Me.chkCONVERTTIME.Checked Then
                    var1a = VerboseNumber(CDec(var1), False)
                Else
                    var1a = var1
                End If
            Else
                var1a = var1
            End If
        End If

        If Me.chkCONVERTTEMP.Checked Then
            str2 = Replace(var3, "deg C", ChrW(176) & "C", 1, -1, CompareMethod.Text)
            str3 = Replace(str2, "degC", ChrW(176) & "C", 1, -1, CompareMethod.Text)
            str3 = Replace(str3, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
        Else
            str3 = Replace(var3, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
        End If

        '20181127 LEE:
        'start doing nbh again
        str3 = Replace(var3, "-", NBHReal, 1, -1, CompareMethod.Text) 'non breaking hyphen


        If Len(strC) = 0 Then
            If Len(var1a) = 0 Or Len(var2) = 0 Then
                str1 = Trim(CStr(var3))
            Else
                str1 = Trim(CStr(var1a) & " " & CStr(var2))
                str1 = Replace(str1, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
                str1 = str1 & " at " & str3
            End If
        Else
            If Len(var1a) = 0 And Len(var2) = 0 Then
                str1 = Trim(str3)
            Else
                'str1 = Trim(CStr(var1a) & " " & CStr(var2) & " at " & CStr(var3))
                str1 = Trim(CStr(var1a) & " " & CStr(var2))
                str1 = Replace(str1, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space
                str1 = str1 & " at " & str3
            End If
        End If
        'If Me.chkCONVERTTEMP.Checked Then
        '    str2 = Replace(str1, "deg C", ChrW(176) & "C", 1, -1, CompareMethod.Text)
        '    str3 = Replace(str2, "degC", ChrW(176) & "C", 1, -1, CompareMethod.Text)
        'Else
        '    str3 = str1
        'End If
        strTP = str1 ' Replace(str3, " ", ChrW(160), 1, -1, CompareMethod.Text) 'non breaking space

        If Len(strC) = 0 Then
            str1 = strTP
        Else
            str1 = strC & " for " & strTP
        End If
        str1 = Trim(str1)

        Me.CHARSTABILITYPERIOD.Text = str1

        'enter charstability period
        dv(intRow).BeginEdit()
        dv(intRow).Item("CHARSTABILITYPERIOD") = str1
        dv(intRow).EndEdit()

        'dgv("CHARSTABILITYPERIOD", intRow).Value = str1

        If Len(str1) = 0 Then

        Else

            strT = Me.txtTitle.Text

            If boolFromInsert Then

                strF = "[#Cycles]"
                If InStr(1, strT, strF, CompareMethod.Text) = 0 And Me.INTNUMBEROFCYCLES.Visible Then
                    boolNoC = True
                Else
                    boolNoC = False
                End If
                strFC = strF
                str2 = Replace(strT, strF, strC, 1, -1, CompareMethod.Text)
                Me.txtTitle.Text = str2
                strT = str2


                strF = "[Period Temp]"
                If InStr(1, strT, strF, CompareMethod.Text) = 0 Then
                    boolNoTP = True
                Else
                    boolNoTP = False
                End If
                strFTP = strF
                str2 = Replace(strT, strF, strTP, 1, -1, CompareMethod.Text)
                Me.txtTitle.Text = str2

                dv(intRow).BeginEdit()
                dv(intRow).Item("CHARHEADINGTEXT") = str2
                dv(intRow).EndEdit()
            End If

        End If

        If boolFromInsert Then
            If boolNoC Then
                strM = "The field code """ & strFC & """ could not be found in the Table Title text box." & ChrW(10) & ChrW(10)
                strM = strM & "If you wish to enter this text into the Table Title, you must copy and paste it into the appropriate position in the Table Title text box."
                MsgBox(strM, MsgBoxStyle.Information, "Couldn't find """ & strF & """")
                Me.CHARSTABILITYPERIOD.Focus()
            End If
            If boolNoTP Then
                strM = "The field code """ & strFTP & """ could not be found in the Table Title text box." & ChrW(10) & ChrW(10)
                strM = strM & "If you wish to enter this text into the Table Title, you must copy and paste it into the appropriate position in the Table Title text box."
                MsgBox(strM, MsgBoxStyle.Information, "Couldn't find """ & strF & """")
                Me.CHARSTABILITYPERIOD.Focus()
            End If
        End If

end1:
        'If boolC Then
        'Else
        '    strM = "The text boxes shown in this group box cannot be blank in order to build a text fragment." & ChrW(10) & ChrW(10)
        '    strM = strM & "If you wish to enter this text into the Table Title, you must copy and paste it into the appropriate position in the Table Title text box."
        '    MsgBox(strM, MsgBoxStyle.Information, "Couldn't find """ & strF & """")
        '    Me.CHARSTABILITYPERIOD.Focus()

        'End If

        '20181112 LEE:
        Me.lblRemember.Visible = False

    End Sub

    Private Sub cmdBuild_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBuild.Click

        Call BuildPeriodTemp(False)

    End Sub

    Function CheckStabilityValidation() As Boolean

        Dim var1
        Dim strM As String
        Dim boolM As Boolean
        Dim ctl As Control

        CheckStabilityValidation = False

        If Me.panCycles.Visible Then
            boolM = False
            var1 = Me.INTNUMBEROFCYCLES.Text
            ctl = Me.INTNUMBEROFCYCLES
            strM = "Entry must integer > to 0"
            If IsNumeric(var1) Then
                If var1 < 1 Then
                    boolM = True
                    GoTo end1
                End If
                If IsInt(var1) Then
                Else
                    boolM = True
                    GoTo end1
                End If
            Else
                boolM = True
                GoTo end1
            End If

        Else

            boolM = False
            var1 = Me.CHARTIMEPERIOD.Text
            ctl = Me.CHARTIMEPERIOD
            strM = "Entry must integer >= to 0"
            If IsNumeric(var1) Then
                If var1 < 0 Then
                    boolM = True
                    GoTo end1
                End If
                If IsInt(var1) Then
                Else
                    boolM = True
                    GoTo end1
                End If
            Else
                boolM = True
                GoTo end1
            End If

            'limited to 9999
            If var1 > 9999 Then
                strM = "Maximum number allowed is 9999"
                boolM = True
                GoTo end1
            End If


            boolM = False
            var1 = Me.CHARTIMEFRAME.Text
            ctl = Me.CHARTIMEFRAME
            strM = "Entry cannot be blank"
            If Len(var1) = 0 Then
                boolM = True
                GoTo end1
            End If

            boolM = False
            var1 = Me.CHARPERIODTEMP.Text
            ctl = Me.CHARPERIODTEMP
            strM = "Entry cannot be blank"
            If Len(var1) = 0 Then
                boolM = True
                GoTo end1
            End If

        End If

end1:
        If boolM Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            ctl.Focus()
            CheckStabilityValidation = False
        Else
            CheckStabilityValidation = True
        End If

    End Function

    Sub GetPos()
        Me.Cursor = New Cursor(Cursor.Current.Handle)
        'MsgBox(Cursor.Position.X & ", " & Cursor.Position.Y)
    End Sub

    Private Sub txtTitle_MouseDown(sender As Object, e As MouseEventArgs) Handles txtTitle.MouseDown

        Dim v

        v = e.Button.ToString

        If StrComp(v, "Right", CompareMethod.Text) = 0 Then

            Call OpenFieldCodes()

            Cursor = Cursors.Default

        End If

    End Sub

    Sub OpenFieldCodes()

        Dim pos As Short
        Dim posS
        Dim posL

        Dim strT As String
        Dim str1 As String
        Dim strL As String
        Dim strR As String


        'record position of cursor in text box
        posS = Me.txtTitle.SelectionStart
        posL = Me.txtTitle.SelectionLength

        pos = posS + posL

        'Me.txtTitle.SelectionLength = 0
        'Me.txtTitle.SelectionStart = pos

        Dim frm As New frmFieldCodes
        frm.gbCopyAll.Visible = False
        'Dim t, l, w, h, t1, l1

        't = Me.txtTitle.Top
        'l = Me.txtTitle.Left
        'w = Me.txtTitle.Width
        'h = Me.txtTitle.Height

        't1 = t + h + Me.Top + 10
        'l1 = l + (w / 2)

        Me.Cursor = New Cursor(Cursor.Current.Handle)

        'frm.Location = new system.drawing.point(l1, t1)

        frm.Location = New System.Drawing.Point(Cursor.Position.X, Cursor.Position.Y + 10)

        frm.ShowDialog()

        If frm.boolCancel Then

            Me.txtTitle.SelectionStart = pos
            Me.txtTitle.SelectionLength = 0

        Else
            strT = Me.txtTitle.Text

            If pos = 0 Then
                strL = "" 'Mid(strT, 1, pos)
                strR = strT 'Mid(strT, pos + 1, Len(strT) - pos)
                'str1 = frm.strFC & " " & strR
                str1 = frm.strFC & strR
            ElseIf pos = Len(strT) Then
                strL = strT 'Mid(strT, 1, pos)
                strR = "" 'Mid(strT, pos + 1, Len(strT) - pos)
                'str1 = strL & " " & frm.strFC
                str1 = strL & frm.strFC

            Else
                strL = Mid(strT, 1, pos)
                strR = Mid(strT, pos + 1, Len(strT) - pos)
                'str1 = strL & " " & frm.strFC & " " & strR
                str1 = strL & frm.strFC & strR
            End If

            'strL = Mid(strT, 1, pos - 1)
            'strR = Mid(strT, pos + 1, Len(strT) - pos)
            'str1 = strL & " " & frm.strFC & " " & strR
            str1 = strL & frm.strFC & strR
            Me.txtTitle.Text = str1

            'select field code

            Me.txtTitle.SelectionStart = posS
            Dim l = Len(frm.strFC)
            Me.txtTitle.SelectionLength = l

        End If

        frm.Dispose()

        Cursor = Cursors.Default

    End Sub

    Private Sub cmsFieldCodes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmsFieldCodes.Click

        Call OpenFieldCodes()

    End Sub


    Private Sub chkDiff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDiff.CheckedChanged

        Dim bool As Boolean
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id As Short
        Dim strM As String

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        bool = Me.chkDiff.Checked


        If bool Then
            dgv = Me.dgvReportTables
            intRow = dgv.CurrentRow.Index
            id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

            Dim bool1 As Boolean
            If boolAllowAcc(id) Then
            Else
                strM = "This table does not support the %Diff parameter."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkDiff.Checked = False
            End If
        End If

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkDiff, "BOOLSTATSDIFF")
            Call CheckStats(Me.chkDiff)
        End If



    End Sub

    Private Sub chkDiffCol_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDiffCol.CheckedChanged

        Dim bool As Boolean
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id As Short
        Dim strM As String

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        bool = Me.chkDiffCol.Checked


        If bool Then
            dgv = Me.dgvReportTables
            intRow = dgv.CurrentRow.Index
            id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

            Dim bool1 As Boolean
            If boolAllowAcc(id) Then
            Else
                strM = "This table does not support the %Diff Column parameter."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkDiffCol.Checked = False
            End If
        End If

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkDiffCol, "BOOLSTATSDIFFCOL")
            Call CheckStats(Me.chkDiffCol)
        End If

    End Sub

    Private Sub chkRegr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRegr.CheckedChanged

        If boolHold Then
        Else

            Dim str1 As String
            Dim str2 As String
            Dim str3 As String

            str1 = Me.txtTitle.Text
            str2 = " and Calibration Curve Parameters"
            If Me.chkRegr.Checked Then 'add verbiage to title
                str3 = str1 & str2
            Else 'remove verbiage to title
                str3 = Replace(str1, str2, "", 1, -1, CompareMethod.Text)
            End If
            'Me.txtTitle.Text = str3

            Call UpdateChkData(False, Me.chkRegr, "BOOLSTATSREGR")
            Call CheckStats(Me.chkRegr)
        End If

    End Sub

    Private Sub rbNR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNR.CheckedChanged

        If boolHold Then
        Else
            'Call UpdateChkData(True, Me.rbNR, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbOutier_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOutier.CheckedChanged

        If boolHold Then
        Else
            'Call UpdateChkData(True, Me.rbOutier, "BOOLSTATSLETTER")
        End If

    End Sub

    Private Sub chkTheoretical_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTheoretical.CheckedChanged

        Dim bool As Boolean
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id As Short
        Dim strM As String

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        bool = Me.chkTheoretical.Checked


        If bool Then
            dgv = Me.dgvReportTables
            intRow = dgv.CurrentRow.Index
            id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

            Dim bool1 As Boolean
            If boolAllowAcc(id) Then
            Else
                strM = "This table does not support the %Theoretical parameter."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkTheoretical.Checked = False
            End If
        End If

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkTheoretical, "BOOLTHEORETICAL")
            Call CheckStats(Me.chkTheoretical)
        End If

    End Sub

    Private Sub chkIncludeAnova_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeAnova.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIncludeAnova, "BOOLINCLANOVA")

        End If
    End Sub

    Private Sub chkBQLLEGEND_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBQLLEGEND.CheckedChanged

        Call UpdateChkData(False, Me.chkBQLLEGEND, "BOOLBQLLEGEND")

    End Sub

    Private Sub cbxSampleG1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleG1.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleG1)

    End Sub

    Private Sub cbxSampleG2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleG2.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleG2)

    End Sub

    Private Sub cbxSampleG3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleG3.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleG3)

    End Sub

    Private Sub cbxSampleG4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleG4.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleG4)

    End Sub

    Private Sub cbxSampleGAD1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleGAD1.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleGAD1)

    End Sub

    Private Sub cbxSampleGAD2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleGAD2.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleGAD2)

    End Sub

    Private Sub cbxSampleGAD3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleGAD3.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleGAD3)

    End Sub

    Private Sub cbxSampleGAD4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleGAD4.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleGAD4)

    End Sub

    Private Sub cbxSampleS1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleS1.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleS1)

    End Sub

    Private Sub cbxSampleS2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleS2.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleS2)

    End Sub

    Private Sub cbxSampleS3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleS3.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleS3)

    End Sub

    Private Sub cbxSampleS4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleS4.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleS4)

    End Sub

    Private Sub cbxSampleSAD1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleSAD1.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleSAD1)

    End Sub

    Private Sub cbxSampleSAD2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleSAD2.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleSAD2)

    End Sub

    Private Sub cbxSampleSAD3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleSAD3.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleSAD3)

    End Sub

    Private Sub cbxSampleSAD4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSampleSAD4.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleSAD4)

    End Sub

    Private Sub chkIncludePSAE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludePSAE.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIncludePSAE, "boolIncludePSAE")
            Call UpdateAnalRun()
        End If

    End Sub

    Sub UpdateAnalRun()

        Dim dgv As DataGridView
        Dim idT As Int64
        Dim intRow As Short

        dgv = Me.dgvReportTables
        If dgv.CurrentRow Is Nothing Then
            Exit Sub
        End If

        'intRow = dgv.CurrentRow.Index
        'idT = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        'If idT = 1 Then
        '    If Me.chkIncludePSAE.Checked Then
        '        frmH.rbAnalRunsShowAll.Checked = True
        '    Else
        '        frmH.rbAnalRunsExclPSAE.Checked = True
        '    End If
        '    Call FillAnalRunSum()
        'End If

    End Sub

    Sub CheckPSAE()

        Dim dv As System.Data.DataView
        Dim intRow As Short
        Dim var1
        Dim intRSA As Short
        Dim idT As Int64

        Dim boolShow As Boolean

        boolShow = False
        Try

            dv = Me.dgvReportTables.DataSource
            intRow = Me.dgvReportTables.CurrentRow.Index

            var1 = dv(intRow).Item("BOOLREQUIRESSAMPLEASSIGNMENT")
            intRSA = NZ(var1, 0)
            var1 = dv(intRow).Item("ID_TBLCONFIGREPORTTABLES")
            idT = NZ(var1, 0)

            'idT Legend
            '1: Summary of Analytical Runs
            '2: Summary of Regression Constants
            '3: Summary of Back-Calculated Calibration Std Conc
            '4: Summary of Interpolated QC Std Conc

            If idT = 2 Or idT = 3 Or idT = 4 Then

            End If



        Catch ex As Exception

        End Try

        If boolShow Then

        End If

    End Sub

    Private Sub rbConc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbConc.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbConc, "BOOLRCCONC")
        End If

    End Sub

    Private Sub rbUsePeakArea_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUsePeakArea.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbUsePeakArea, "BOOLRCPA")
        End If

    End Sub

    Private Sub rbUseISPeakArea_CheckedChanged(sender As Object, e As EventArgs) Handles rbUseISPeakArea.CheckedChanged

        '20190225 LEE:
        'Logic: push rbUsePeakArea
        'In table generation code, if BOOLRCPA, BOOLRCCONC, and BOOLRCPARATIO are all false, then use IS
        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbUsePeakArea, "BOOLRCPA")
        End If

    End Sub

    Private Sub rbUsePeakAreaRatio_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUsePeakAreaRatio.CheckedChanged

        Call PARChange(True)

    End Sub

    Sub PARChange(boolMsg As Boolean)

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbUsePeakAreaRatio, "BOOLRCPARATIO")
        End If

        Dim boolPAR As Boolean = Me.rbUsePeakAreaRatio.Checked

        If boolTS Then
        Else

            Try
                Dim dgv As DataGridView = Me.dgvReportTables
                Dim intID As Short
                Dim intIndex As Short
                intIndex = dgv.CurrentRow.Index
                intID = dgv("ID_TBLCONFIGREPORTTABLES", intIndex).Value
                Dim strM As String
                Dim strM1 As String

                strM = "If 'Peak Area Ratio' or 'Show Table as Matrix Factor' is checked, 'Include Int Std Table' is not availabe."

                Select Case intID
                    Case 22
                        Me.chkIncludeIS.Enabled = True
                        Me.chkCustomLeg.Enabled = True
                        GoTo end1

                    Case 17

                    Case Else

                        If boolPAR And Me.panIS.Visible And boolMsg And Me.chkIncludeIS.Checked Then

                            MsgBox(strM, vbInformation, "FYI...")

                        End If
                End Select

              

            Catch ex As Exception
                Dim var1
                var1 = ex.Message
            End Try

        End If

        If boolPAR Then
         
        Else
            Me.chkIncludeIS.Enabled = True
            Me.chkCustomLeg.Enabled = True
        End If

end1:

    End Sub

    Private Sub chkIncludeIS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeIS.CheckedChanged


        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIncludeIS, "BOOLINCLUDEISTBL")

        End If

    End Sub


    Private Sub rbPosLeg_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPosLeg.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbPosLeg, "BOOLPOSLEG")
            Call DoLegendThings()
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")
        End If

    End Sub

    Private Sub chkCustomLeg_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomLeg.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkCustomLeg, "BOOLCUSTOMLEG")
            'Call DoLegendThings()
        End If

    End Sub

    Private Sub rbNegLeg_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNegLeg.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbNegLeg, "BOOLNEGLEG")
            Call DoLegendThings()
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")
        End If

    End Sub

    Private Sub CHARTITLELEG_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARTITLELEG.Validated

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
        End If

    End Sub

    Private Sub CHARNUMLEG_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARNUMLEG.Validated

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
        End If

    End Sub

    Private Sub CHARDENLEG_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARDENLEG.Validated

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")
        End If

    End Sub

    Private Sub chkNoneLeg_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNoneLeg.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkNoneLeg, "BOOLNONELEG")
            'Call DoLegendThings()
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")

        End If

    End Sub

    Private Sub chkIncludeDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeDate.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIncludeDate, "BOOLINCLUDEDATE")

        End If

    End Sub

    Private Sub rbDifference_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDifference.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbDifference, "BOOLDIFFERENCE")
            Call DoLegendThings()
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")
            Call UpdateNomDenom()
        End If

    End Sub

    Private Sub rbRecovery_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRecovery.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbRecovery, "BOOLRECOVERY")
            Call DoLegendThings()
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")
            Call UpdateNomDenom()
        End If

    End Sub

    Private Sub rbMeanAcc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMeanAcc.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbMeanAcc, "BOOLMEANACCURACY")
            Call DoLegendThings()
            Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
            Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
            Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")
            Call UpdateNomDenom()
        End If

    End Sub

    Sub UpdateNomDenom()

        Dim strNum As String
        Dim strDenom As String

        Select Case gidCRT

            Case 13, 14, 15

            Case Else
                If BOOLDIFFERENCE Then
                    Me.gbNumerator.Visible = True
                    Me.gbDenom.Visible = True
                ElseIf boolRECOVERY Then
                    Me.gbNumerator.Visible = True
                    Me.gbDenom.Visible = True
                ElseIf boolMEANACCURACY Then
                    Me.gbNumerator.Visible = True
                    Me.gbDenom.Visible = False
                End If

                If BOOLCALCINTSTDNMF Then
                    Me.rbOld.Checked = True
                    Me.rbNew.Checked = False
                Else
                    Me.rbOld.Checked = False
                    Me.rbNew.Checked = True
                End If
        End Select

    End Sub

    Private Sub chkIncludeWatsonLabel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeWatsonLabel.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIncludeWatsonLabel, "BOOLINCLUDEWATSONLABELS")

        End If

    End Sub


    Private Sub chkIncludeAnovaSumStats_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeAnovaSumStats.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIncludeAnovaSumStats, "BOOLINCLANOVASUMSTATS")

        End If
    End Sub

    Private Sub chkRE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRE.CheckedChanged

        Dim bool As Boolean
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim id As Short
        Dim strM As String

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        bool = Me.chkRE.Checked


        If bool Then
            dgv = Me.dgvReportTables
            intRow = dgv.CurrentRow.Index
            id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

            Dim bool1 As Boolean
            If boolAllowAcc(id) Then
            Else
                strM = "This table does not support the %RE parameter."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.chkRE.Checked = False
            End If
        End If

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkRE, "BOOLSTATSRE")
            Call CheckStats(Me.chkRE)
        End If

    End Sub


    Private Sub CHARSTABILITYPERIOD_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARSTABILITYPERIOD.Validated

        Call UpdatecharStabilityPeriod()

    End Sub

    Private Sub cmdInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsert.Click

        Call BuildPeriodTemp(True)

    End Sub


    Private Sub INTNUMBEROFCYCLES_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles INTNUMBEROFCYCLES.Validated

        If boolHold Then
        Else
            Call UpdateTBData(Me.INTNUMBEROFCYCLES, "INTNUMBEROFCYCLES")
        End If

    End Sub

    Private Sub CHARTIMEPERIOD_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARTIMEPERIOD.Validated

        Dim var1, var2

        var1 = Me.CHARTIMEPERIOD.Text
        If Me.chkCONVERTTIME.Checked And Len(var1) <> 0 And IsNumeric(var1) Then
            var2 = VerboseNumber(CDec(var1), False)
            Me.CHARTIMEPERIOD.Text = var2
        End If

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARTIMEPERIOD, "CHARTIMEPERIOD")
        End If

    End Sub

    Private Sub CHARTIMEFRAME_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARTIMEFRAME.Validated

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARTIMEFRAME, "CHARTIMEFRAME")
        End If

    End Sub


    Private Sub CHARPERIODTEMP_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARPERIODTEMP.Validated

        Dim var1, var2
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        var1 = Me.CHARPERIODTEMP.Text
        If Me.chkCONVERTTEMP.Checked And Len(var1) <> 0 Then
            str1 = CStr(var1)
            str2 = Replace(str1, "deg C", ChrW(176) & "C", 1, -1, CompareMethod.Text)
            str3 = Replace(str2, "degC", ChrW(176) & "C", 1, -1, CompareMethod.Text)
            Me.CHARPERIODTEMP.Text = str3
        End If

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARPERIODTEMP, "CHARPERIODTEMP")
        End If

    End Sub

    Private Sub chkCONVERTTIME_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCONVERTTIME.CheckedChanged

        Dim var1

        var1 = Me.CHARTIMEPERIOD.Text

        Call ConvertTime(var1, True, Me.chkCONVERTTIME.Checked)

    End Sub

    Function ConvertTime(var1 As Object, boolDoUpdate As Boolean, boolChecked As Boolean) As String

        Dim var2

        ConvertTime = NZ(var1, "")

        If boolChecked And Len(var1) <> 0 And IsNumeric(var1) Then
            var2 = VerboseNumber(CDec(var1), False)
            Me.CHARTIMEPERIOD.Text = var2
            ConvertTime = NZ(var2, "")
        End If

        If boolHold Or boolDoUpdate = False Then
        Else
            Call UpdateChkData(False, Me.chkCONVERTTIME, "BOOLCONVERTTIME")
        End If

    End Function

    Private Sub chkCONVERTTEMP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCONVERTTEMP.CheckedChanged

        Dim var1

        var1 = Me.CHARPERIODTEMP.Text

        Call ConvertTemp(var1, True, Me.chkCONVERTTEMP.Checked)

    End Sub

    Function ConvertTemp(var1 As Object, boolDoUpdate As Boolean, boolChecked As Boolean) As String

        Dim var2
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        ConvertTemp = var1

        If boolChecked And Len(var1) <> 0 Then
            str1 = CStr(var1)
            str2 = Replace(str1, "deg C", ChrW(176) & "C", 1, -1, CompareMethod.Text)
            str3 = Replace(str2, "degC", ChrW(176) & "C", 1, -1, CompareMethod.Text)
            Me.CHARPERIODTEMP.Text = str3
            ConvertTemp = str3
        End If

        If boolHold Or boolDoUpdate = False Then
        Else
            Call UpdateChkData(False, Me.chkCONVERTTEMP, "BOOLCONVERTTEMP")
        End If

    End Function

    Private Sub INTNUMBEROFCYCLES_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles INTNUMBEROFCYCLES.Validating

        Dim strM As String
        Dim var1
        Dim boolE As Boolean

        strM = "Entry must be integer >= 0"
        var1 = Me.INTNUMBEROFCYCLES.Text
        boolE = False

        var1 = NZ(var1, "")

        If Len(var1) = 0 Then
            boolE = False
            GoTo end1
        End If

        If IsNumeric(var1) Then
        Else
            boolE = True
            GoTo end1
        End If

        If var1 >= 0 Then
        Else
            boolE = True
            GoTo end1
        End If

end1:
        If boolE Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
        End If

    End Sub

    Private Sub CHARISCONC_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHARISCONC.Validated
        Dim var1, var2
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARISCONC, "CHARISCONC")
        End If
    End Sub

    Private Sub frmReportTableConfig_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged

        If Me.Visible = False Then

            boolClose = True

            'make sure selected row is visible
            Try
                Dim intRow As Short
                Dim dgv As DataGridView

                dgv = frmH.dgvReportTableConfiguration
                intRow = dgv.CurrentRow.Index
                dgv.CurrentCell = dgv("INTORDER", intRow)


            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub chkIntraRunSumStats_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIntraRunSumStats.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIntraRunSumStats, "BOOLINTRARUNSUMSTATS")

        End If

    End Sub

    Private Sub CHARSTABILITYPERIOD_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CHARSTABILITYPERIOD.Validating


        Try
            Dim strM As String
            Dim str1 As String

            Dim dgv As DataGridView = Me.dgvReportTables
            Dim int1 As Short = dgv.CurrentCell.RowIndex
            Dim str2 As String = dgv("CHARHEADINGTEXT", int1).Value

            Dim strMod As String = "Advanced Report Table Configuration - " & str2
            Dim strSource As String = "Stability Conditions Summary"


            str1 = Me.CHARSTABILITYPERIOD.Text
            If CheckColLenEx(str1, 255, strMod, strSource) Then
                e.Cancel = True
            End If
        Catch ex As Exception

        End Try
       

    End Sub

    Private Sub CHARTITLELEG_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CHARTITLELEG.Validating

        Try
            Dim strM As String
            Dim str1 As String

            Dim dgv As DataGridView = Me.dgvReportTables
            Dim int1 As Short = dgv.CurrentCell.RowIndex
            Dim str2 As String = dgv("CHARHEADINGTEXT", int1).Value

            Dim strMod As String = "Advanced Report Table Configuration - " & str2
            Dim strSource As String = "Title Legends - Title"

            str1 = Me.CHARTITLELEG.Text
            If CheckColLenEx(str1, 250, strMod, strSource) Then
                e.Cancel = True
            End If
        Catch ex As Exception

        End Try
      

    End Sub

    Private Sub CHARNUMLEG_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CHARNUMLEG.Validating

        Dim strM As String
        Dim str1 As String

        Dim dgv As DataGridView = Me.dgvReportTables
        Dim int1 As Short = dgv.CurrentCell.RowIndex
        Dim str2 As String = dgv("CHARHEADINGTEXT", int1).Value

        Dim strMod As String = "Advanced Report Table Configuration - " & str2
        Dim strSource As String = "Title Legendes - Numerator"

        str1 = Me.CHARNUMLEG.Text
        If CheckColLenEx(str1, 250, strMod, strSource) Then
            e.Cancel = True
        End If


    End Sub

    Private Sub CHARDENLEG_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHARDENLEG.TextChanged

    End Sub

    Private Sub CHARDENLEG_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CHARDENLEG.Validating

        Dim strM As String
        Dim str1 As String

        str1 = Me.CHARDENLEG.Text

        Dim dgv As DataGridView = Me.dgvReportTables
        Dim int1 As Short = dgv.CurrentCell.RowIndex
        Dim str2 As String = dgv("CHARHEADINGTEXT", int1).Value

        Dim strMod As String = "Advanced Report Table Configuration - " & str2
        Dim strSource As String = "Title Legendes - Denominator"

        If CheckColLenEx(str1, 250, strMod, strSource) Then
            e.Cancel = True
        End If

    End Sub

    Private Sub dgvAnalytes_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvAnalytes.CellValidating

        Dim dgv As DataGridView

        dgv = Me.dgvAnalytes
        Dim intRow As Short
        Dim intCol As Short
        Dim strCol As String
        Dim strM As String
        Dim var1, var2, var3
        Dim boolE As Boolean = False

        strM = "Entry must be numeric > 0"

        intRow = e.RowIndex
        intCol = e.ColumnIndex

        strCol = dgv.Columns(intCol).Name

        Select Case strCol
            Case "NUMINCSAMPLECRIT01"
                'entry must be number > 0
                var1 = e.FormattedValue
                If Len(var1) = 0 Then
                    boolE = True
                    GoTo end1
                End If

                If IsNumeric(var1) Then
                Else
                    boolE = True
                    GoTo end1
                End If

                If var1 < 0 Then
                    boolE = True
                    GoTo end1
                End If

                'enter value in tblReportTableAnalytes
                Dim Count1 As Short
                Dim strF As String
                Dim str1 As String
                Dim id1 As Int64

                var2 = dgv("ID_TBLREPORTTABLEANALYTES", intRow).Value

                strF = "ID_TBLREPORTTABLEANALYTES = " & var2

                Dim rows() As DataRow = tblReportTableAnalytes.Select(strF)
                rows(0).BeginEdit()
                rows(0).Item("NUMINCSAMPLECRIT01") = var1
                rows(0).EndEdit()
        End Select

end1:

        If boolE Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
        End If

    End Sub


    ' Nick Addition
    Private Sub SetToEditMode()

        Me.cmdEdit.Enabled = False
        Me.cmdEdit.BackColor = System.Drawing.Color.Gray
        Me.cmdSave.Enabled = True
        Me.cmdSave.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancel.Enabled = True
        Me.cmdCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.Enabled = False
        Me.cmdExit.BackColor = System.Drawing.Color.Gray

        Me.cmdPasteConditions.Enabled = True
        Me.cmdPasteConditions.BackColor = System.Drawing.Color.Gainsboro

    End Sub

    Private Sub SetToNonEditMode()

        Me.cmdEdit.Enabled = True
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSave.Enabled = False
        Me.cmdSave.BackColor = System.Drawing.Color.Gray
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray
        Me.cmdExit.Enabled = True
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro

        Me.cmdPasteConditions.Enabled = False
        Me.cmdPasteConditions.BackColor = System.Drawing.Color.Gray

    End Sub
    Private Sub frmReportTableConfig_ToolTipSet()

        Dim idC As Int32 = idCGet()


        Dim str1 As String

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()

        ' Set up the delays for the ToolTip.
        'toolTip1.AutoPopDelay = 5000
        'toolTip1.InitialDelay = 250
        'toolTip1.ReshowDelay = 50

        toolTip1.AutomaticDelay = intToolTipDelay
        'toolTip1.UseFading = False
        'tooltip1.
        'toolTip1.BackColor = Color.Goldenrod
        'toolTip1.IsBalloon = True
        ' Force the ToolTip text to be displayed whether or not the form is active.
        toolTip1.ShowAlways = True

        Try

            'Set mode buttons
            toolTip1.SetToolTip(Me.chkSD, "Report Standard Deviation")
            toolTip1.SetToolTip(Me.chkCV, "Report % Coefficient of Variation = [(Standard Deviation)/Mean] * 100")
            toolTip1.SetToolTip(Me.chkBias, "Report Average of '% Bias' = [[Mean/(Nominal Concentration)]-1]*100" & vbCrLf &
                                "Note: %Relative Error = %Diff = %Bias")
            toolTip1.SetToolTip(Me.chkDiff, "Report Average of '% Difference' = [[Mean/(Nominal Concentration)]-1]*100" & vbCrLf &
                                "Note: %Relative Error = %Diff = %Bias")
            toolTip1.SetToolTip(Me.chkRE, "Report Average of '% Relative Error' = [[Mean/(Nominal Concentration)]-1]*100" & vbCrLf &
                                "Note: %Relative Error = %Diff = %Bias")
            toolTip1.SetToolTip(Me.chkDiffCol, "Add extra report column with '% Difference' for each concentration, where " & vbCrLf &
                                "% Difference = [[(Measured Concentration) - (Nominal Concentration)] / (Nominal Concentration)] * 100")
            toolTip1.SetToolTip(Me.chkTheoretical, "Report theoretical concentration percentage = [Mean/(Nominal Concentration)] * 100")

            toolTip1.SetToolTip(Me.chkN, "Report number of samples (n)")
            toolTip1.SetToolTip(Me.chkIncludeDate, "Report the Start Analysis Date for each run.  (The date format is configured in the " & vbCrLf &
                                """Add/Edit Top Level Data: Study Configuration"" area in StudyDoc.)")
            toolTip1.SetToolTip(Me.rbConc, "Report (and perform any comparisons) on concentration values (requires regression)")
            toolTip1.SetToolTip(Me.rbUsePeakArea, "Report (and perform any comparisons) on peak areas")
            toolTip1.SetToolTip(Me.rbUsePeakAreaRatio, "Report (and perform any comparisons) on peak area ratios [(Analyte Peak Area) / (Internal Standard Peak Area)]")

            toolTip1.SetToolTip(Me.chkIncludeIS, "Include a separate table with the Internal Standard Peak Areas.")
            toolTip1.SetToolTip(Me.lblCHARISCONC, "Concentration and Units (e.g. 5 ng/ml) of Internal Standard Nominal Concentration")

            toolTip1.SetToolTip(Me.rbDifference, "Report the % difference between means. Calcuation details for current selection" & vbCrLf &
                                "shown below. (to see, ensure ""Custom Legend"" is not checked)")
            toolTip1.SetToolTip(Me.rbRecovery, "Report ratio of means. Calcuation details for current selection" & vbCrLf &
                                "shown below. (to see, ensure ""Custom Legend"" is not checked)")
            toolTip1.SetToolTip(Me.rbMeanAcc, "Report % difference between means. Calcuation details for current selection" & vbCrLf &
                                "shown below. (to see, ensure ""Custom Legend"" is not checked)")

            toolTip1.SetToolTip(Me.rbPosLeg, "Use ""Old – New""  or ""Old/New"" in the Calculation." & vbCrLf &
                                "(Note: For Ad Hoc Table, ""Old"" group is group containing lowest run/sequence sample)")

            toolTip1.SetToolTip(Me.rbNegLeg, "Use ""New – Old""  or ""New/Old"" in the Calculation." & vbCrLf &
                                "(Note: For Ad Hoc Table, ""Old"" group is group containing lowest run/sequence sample)")

            toolTip1.SetToolTip(Me.lblCHARTITLELEG, "Title for Legend (e.g. ""% Difference = "".  This is placed in the legend in the form:  Title   Numer / Denom")
            toolTip1.SetToolTip(Me.CHARTITLELEG, "Title for Legend (e.g. ""% Difference = "".  This is placed in the legend in the form:  Title   Numer / Denom")
            toolTip1.SetToolTip(Me.lblCHARNUMLEG, "Numerator for Legend (e.g. ""(Mean New – Mean Old) * 100"".  This is placed in the legend in the form:  Title    Numer / Denom")
            toolTip1.SetToolTip(Me.CHARNUMLEG, "Numerator for Legend (e.g. ""(Mean New – Mean Old) * 100"".  This is placed in the legend in the form:  Title    Numer / Denom")
            toolTip1.SetToolTip(Me.lblCHARDENLEG, "Denominator for Legend (e.g. ""(Mean Old)"".  This is placed in the legend in the form:  Title    Numer / Denom")
            toolTip1.SetToolTip(Me.CHARDENLEG, "Denominator for Legend (e.g. ""(Mean Old)"".  This is placed in the legend in the form:  Title    Numer / Denom")

            toolTip1.SetToolTip(Me.chkNoneLeg, "Do not include a standard legend showing the comparison calculation. " & vbCrLf &
                                "(Note: A custom legend overrides this option)")
            toolTip1.SetToolTip(Me.chkCustomLeg, "Include only an Int Std table with this selection. ")

            str1 = "Check to include individual calculated Recovery/MatrixFactor values in the Recovery/MatrixFactor column of the table."
            str1 = str1 & ChrW(10) & "Used if it is desired to report statistics for the Recovery/MatrixFactor value."
            toolTip1.SetToolTip(Me.chkBOOLDOINDREC, str1)

            toolTip1.SetToolTip(Me.INTNUMBEROFCYCLES, "Number of freeze/thaw cycles")
            toolTip1.SetToolTip(Me.cmdBuild, "Create text phrase based on data")
            toolTip1.SetToolTip(Me.cmdInsert, "Replace [Period Temp] (and also [#Cycles] if relevant) in table title with values above." & vbCrLf &
                                "(Note: This will also update the Stability Conditions Summary to the right.")
            Me.dgvReportTables.Columns.Item("CHARTABLENAME").ToolTipText = "Choose table to set options for"
            Me.dgvReportTables.Columns.Item("CHARHEADINGTEXT").ToolTipText = "Choose table to set options for"

            str1 = "Check to calculate Internal Standard-Normalized Matrix Factor (MF) = MF(Analyte)/MF(IntStd)."
            str1 = str1 & ChrW(10) & "Uncheck to calculate Internal Standard-Normalized Matrix Factor (MF) = AreaRatio(PES)/AreaRatio(RS)."""
            toolTip1.SetToolTip(Me.chkCalcIntStdNMF, str1)

            Select Case idC
                Case 17
                    str1 = "Check to include a Matrix Factor statistics section."
                Case Else
                    str1 = "Check to report table as a Matrix Factor table rather than a Suppression/Enhancement table."
            End Select

            toolTip1.SetToolTip(Me.chkMFTable, str1)

            Select Case idC
                Case 17
                    str1 = "As opposed to (Average Peak Area Lot) / (Average Peak Area Solvent)."
                Case Else
                    If Me.chkCalcIntStdNMF.Checked Then
                        str1 = "Check to show Matrix Factor columns for Analyte and Int Std."
                        str1 = str1 & ChrW(10) & "Uncheck to show Peak Area Ratio columns for Analyte and Int Std."
                    Else
                        str1 = "Check to include Peak Area Ratio columns Recovery and Post Extraction Spiking solutions."
                    End If
            End Select
       
            toolTip1.SetToolTip(Me.chkInclMFCols, str1)

            str1 = "Check to include an Internal Standard-Normalized column."
            toolTip1.SetToolTip(Me.chkInclIntStdNMF, str1)

            str1 = "Enter how the Carryover Injection is to referred to in the table column labels."
            str1 = str1 & ChrW(10) & "E.g. 'Blank' or 'Carryover Injection'"
            toolTip1.SetToolTip(Me.CHARCARRYOVERLABEL, str1)

            str1 = "Group QC Levels by Watson Assay Level ID."
            toolTip1.SetToolTip(Me.rbINTQCLEVELGROUPLevel, str1)

            str1 = "Group QC Levels Nominal Concentration."
            toolTip1.SetToolTip(Me.rbINTQCLEVELGROUPNomConc, str1)

            str1 = "Group QC Levels by Watson QC Label ID's."
            toolTip1.SetToolTip(Me.rbINTQCLEVELGROUPQCLabel, str1)

            '20190109 LEE:
            'tooltips for stability radiobuttons
            str1 = "Autosampler stability: The sponsor should demonstrate the stability of extracts in the autosampler" & ChrW(10) & "only if the autosampler storage conditions are different or not covered by extract (processed sample) stability."
            toolTip1.SetToolTip(Me.rbAutosampler, str1)

            str1 = "Bench-top stability: The sponsor should determine the stability of samples under the laboratory handling conditions" & ChrW(10) & "that are expected for the study samples (e.g., the stability of samples maintained at room temperature or stored in an ice bucket)."
            toolTip1.SetToolTip(Me.rbBenchTop, str1)

            str1 = "Extract (or processed sample) stability: The sponsor should assess the stability of processed samples," & ChrW(10) & "including the residence time in the autosampler against freshly prepared calibrators."
            toolTip1.SetToolTip(Me.rbProcess, str1)

            str1 = "Freeze-thaw stability: The sponsor should assess the stability of the sample after a minimum of three freeze-thaw cycles." & ChrW(10) & "QC samples should be thawed and analyzed according to the same procedures as the study samples." & ChrW(10) & "QC samples should be frozen for at least 12 hours between cycles." & ChrW(10) & "Freeze-thaw stability QCs should be compared to freshly prepared calibration curves and QCs."
            toolTip1.SetToolTip(Me.rbFT, str1)

            str1 = "Long-term stability: The sponsor should determine the long-term stability of the sample over a period of time" & ChrW(10) & "equal to or exceeding the time between the date of first sample collection and the date of last sample analysis." & ChrW(10) & "The storage temperatures studied should be the same as those used to store study samples." & ChrW(10) & "Long-term stability QCs should be compared to freshly prepared calibration curves and QCs." & ChrW(10) & "Determination of stability at minus 20ºC would cover stability at colder temperatures."
            toolTip1.SetToolTip(Me.rbLT, str1)

            str1 = "Reinjection Stability"
            toolTip1.SetToolTip(Me.rbReinjection, str1)

            str1 = "Batch Reinjection Stability"
            toolTip1.SetToolTip(Me.rbBatchReinjection, str1)

            str1 = "Whole Blood Stability"
            toolTip1.SetToolTip(Me.rbBlood, str1)

            str1 = "Stock solution stability: Stock solutions should not be made from reference materials that are about to expire" & ChrW(10) & "unless the purity of the analyte in the stock solutions is re-established." & ChrW(10) & "When the stock solution exists in a different state (e.g., solution versus solid)" & ChrW(10) & "or in a different buffer composition (which is generally the case for macromolecules)" & ChrW(10) & "from the certified reference standard, the sponsor should generate stability data" & ChrW(10) & "on stock solutions to justify the duration of stock solution storage stability."
            toolTip1.SetToolTip(Me.rbStockSolution, str1)

            str1 = "Spiking solution stability: Stock solutions should not be made from reference materials that are about to expire" & ChrW(10) & "unless the purity of the analyte in the stock solutions is re-established." & ChrW(10) & "When the stock solution exists in a different state (e.g., solution versus solid)" & ChrW(10) & "or in a different buffer composition (which is generally the case for macromolecules)" & ChrW(10) & "from the certified reference standard, the sponsor should generate stability data" & ChrW(10) & "on stock solutions to justify the duration of stock solution storage stability."
            toolTip1.SetToolTip(Me.rbSpiking, str1)

        Catch ex As Exception

        End Try


    End Sub

    Private Sub panFormat_Paint(sender As Object, e As PaintEventArgs) Handles panFormat.Paint

    End Sub


    Private Sub txtTitle_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtTitle.Validating

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim str1 As String
        Dim strM As String

        str1 = Me.txtTitle.Text
        If Len(str1) > 255 Then
            e.Cancel = True

            strM = "This field is limited in length to 255." & ChrW(10) & ChrW(10)
            strM = strM & "The length of the entered text is " & Len(str1) & "." & ChrW(10) & ChrW(10)
            strM = strM & "Please modify the text to conform to the defined text limit."
            strM = strM & ChrW(10) & ChrW(10) & "Advanced Report Table Configuration - Title cell"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")

            GoTo end1

        End If

        If Len(str1) = 0 Then
            strM = "Table Title cannot be blank"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
            GoTo end1
        Else
            dgv = Me.dgvReportTables
            intRow = dgv.CurrentRow.Index
            dgv("CHARHEADINGTEXT", intRow).Value = Me.txtTitle.Text
        End If

end1:

    End Sub


    Sub MakeSASTable()

        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim str1 As String
        Dim str2 As String
        Dim strT As String
        Dim var1, var2
        Dim int1 As Int32
        Dim int2 As Int32



        '1	Summary of Analytical Runs
        '2	Summary of Regression Constants
        '3	Summary of Back-Calculated Calibration Std Conc
        '4	Summary of Interpolated QC Std Conc
        '5	Summary of Samples
        '6	Summary of Reassayed Samples
        '7	Summary of Repeat Samples
        '11	Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
        '12	Summary of Interpolated Dilution QC Concentrations
        '13	Summary of Combined Recovery
        '14	Summary of True Recovery
        '15	Summary of Suppression/Enhancement
        '17	Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
        '18	Summary of [Period Temp] Stability in Matrix
        '19	Summary of Freeze/Thaw [#Cycles] Stability in Matrix
        '21	[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
        '22	[Period Temp] Stock Solution Stability Assessment
        '23	[Period Temp] Spiking Solution Stability Assessment
        '29	[Period Temp] Long-Term QC Std Storage Stability
        '30:     Incurred Samples
        '31	Ad Hoc QC Stability Table
        '32	Ad Hoc QC Stability Comparison Table
        '33	System Suitability Table v1
        '34	Selectivity in Individual Lots Table v1
        '35	Carryover in Individual Lots Table v1
        '36	Method Trial Back-Calculated Calibration Std Conc v1
        '37	Method Trial Control and Fortified QC Samples v1
        '38	Method Trial Incurred Blinded Samples v1


        tblSAS = PivotASP(tblAutoAssignSamples, 0)
        tblSAS.AcceptChanges()

        'debug
        int1 = tblSAS.Rows.Count
        int2 = tblSAS.Columns.Count

        Dim dv As DataView = New DataView(tblSAS, "", "INTORDER ASC", DataViewRowState.CurrentRows)
        Dim dgv As DataGridView = Me.dgvSAS

        dv.AllowDelete = False
        dv.AllowEdit = True
        dv.AllowNew = False

        dgv.DataSource = dv

        'debug
        int1 = dgv.Rows.Count
        int2 = dgv.ColumnCount

        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns(Count1).Visible = False
            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        dgv.RowHeadersWidth = 20


        dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

        dgv.Columns("CHARLABEL").HeaderText = "Text Fragment Type"
        dgv.Columns("CHARLABEL").Visible = True
        dgv.Columns("CHARLABEL").ReadOnly = True
        dgv.Columns("CHARLABEL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        dgv.Columns("CHARLABEL").DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Regular)

        dgv.Columns("CHARVALUE").HeaderText = "Text Fragment" & ChrW(10) & "(1,2,3,5,6)"
        dgv.Columns("CHARVALUE").Visible = True
        dgv.Columns("CHARVALUE").ReadOnly = True
        dgv.Columns("CHARVALUE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        dgv.Columns("CHARVALUE").DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Regular)

        dgv.Columns("CHARNOT").HeaderText = "Exclude" & ChrW(10) & "(4,5)"
        dgv.Columns("CHARNOT").Visible = True
        dgv.Columns("CHARNOT").ReadOnly = True
        dgv.Columns("CHARNOT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        dgv.Columns("CHARNOT").DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Regular)

        dgv.Columns("CHAREXAMPLE").HeaderText = "Example"
        dgv.Columns("CHAREXAMPLE").Visible = True
        dgv.Columns("CHAREXAMPLE").ReadOnly = True
        dgv.Columns("CHAREXAMPLE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
        dgv.Columns("CHAREXAMPLE").DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Regular)

        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns(Count1).ReadOnly = True
        Next

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        dgv.AutoResizeColumns()

        dgv.Columns("CHARVALUE").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgv.Columns("CHARNOT").AutoSizeMode = DataGridViewAutoSizeColumnMode.None




        'now load tblASP to dgvASP
        Dim dvASP As DataView = tblAutoAssignSamples.AsDataView
        Me.dgvASP.DataSource = dvASP

        ''debug
        ''console.writeline("Start")
        'var1 = ""
        'For Count1 = 0 To tblSAS.Columns.Count - 1
        '    var2 = tblSAS.Columns(Count1).ColumnName
        '    var1 = var1 & ";" & var2
        'Next
        ''console.writeline(var1)

        'For Count2 = 0 To tblSAS.Rows.Count - 1
        '    var1 = ""
        '    For Count1 = 0 To tblSAS.Columns.Count - 1
        '        var2 = NZ(tblSAS.Rows(Count2).Item(Count1), "")
        '        var1 = var1 & ";" & var2
        '    Next
        '    'console.writeline(var1)
        'Next
        ''console.writeline("End")

    End Sub


    Sub FilterSAS()

        Dim var1

        Try

            Dim dgv As DataGridView = Me.dgvSAS

            Dim dv As DataView = dgv.DataSource

            Dim dgv1 As DataGridView = Me.dgvASP
            Dim dgv2 As DataGridView = Me.dgvReportTables

            Dim dv1 As DataView = dgv1.DataSource

            Dim id As Int64
            Dim intRow As Integer
            Dim strF1 As String
            Dim strF2 As String
            Dim str1 As String
            Dim str2 As String
            Dim idCT As Int64
            Dim Count1 As Short

            Dim int1 As Short
            Dim int2 As Short

            str2 = "Example Exclude: LLOQ OR Dil"

            intRow = dgv2.CurrentRow.Index

            id = dgv2("ID_TBLREPORTTABLE", intRow).Value
            idCT = dgv2("ID_TBLCONFIGREPORTTABLES", intRow).Value
            strF1 = "(ID_TBLREPORTTABLE = " & id & " AND ID_TBLCONFIGREPORTTABLES = " & idCT & ")"
            str1 = ReturnSASRows(idCT)
            If Len(str1) = 0 Then
                strF2 = "(ID_TBLREPORTTABLE = 0)"
            Else
                strF2 = strF1 & " AND (" & str1 & ")"
            End If

            'Console.Write(strF2)

            'strF2 = strF1

            dv.RowFilter = strF2 ' "ID_TBLREPORTTABLE = " & id
            ''console.writeline(strF2)
            int1 = dv.Count
            dv.AllowDelete = False
            dv.AllowNew = False
            dv.AllowEdit = True

            'do the same for dgvASP
            dv1.RowFilter = "ID_TBLREPORTTABLE = " & id
            dv1.AllowDelete = False
            dv1.AllowNew = False

            dgv.AutoResizeRows()

            Try
                int1 = InStr(1, strF2, "BOOLUSESTDCOLLABELS", CompareMethod.Text)

                If int1 > 0 Then

                    'fill examples
                    str1 = ""
                    Select Case idCT

                        Case 2, 3, 34
                            'str1 = "STD 1, STD 2, etc."
                        Case Else
                            str1 = "QC LLOQ, QC Low, QC Mid (QC Mid-1, QC Mid-2), QC High, QC Diln"
                    End Select

                    dgv("CHAREXAMPLE", 0).Value = str1

                    dgv("CHARVALUE", 0).DataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


                    'Dim chk As New DataGridViewCheckBoxCell
                    'col1.DataType = System.Type.GetType("System.Boolean")
                    'chk.ValueType = System.Type.GetType("System.String")
                    'chk.ValueType = System.Type.GetType("System.Int16")
                    'chk.Value = False

                    '20160808 LEE: Can't use checkbox, keep getting data type error
                    'try combobox instead

                    Dim cbx As New DataGridViewComboBoxCell
                    cbx.Items.Add("TRUE")
                    cbx.Items.Add("FALSE")
                    cbx.DisplayStyleForCurrentCellOnly = True
                    cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
                    Try
                        dgv("CHARVALUE", 0) = cbx
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                    End Try

                End If

                int2 = InStr(1, strF2, "BOOLACCEPTEDONLY", CompareMethod.Text)

                If int2 > 0 Then

                    dgv("CHARVALUE", dgv.Rows.Count - 1).DataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


                    'Dim chk As New DataGridViewCheckBoxCell
                    'col1.DataType = System.Type.GetType("System.Boolean")
                    'chk.ValueType = System.Type.GetType("System.String")
                    'chk.ValueType = System.Type.GetType("System.Int16")
                    'chk.Value = False

                    '20160808 LEE: Can't use checkbox, keep getting data type error
                    'try combobox instead

                    Dim cbx As New DataGridViewComboBoxCell
                    cbx.Items.Add("TRUE")
                    cbx.Items.Add("FALSE")
                    cbx.DisplayStyleForCurrentCellOnly = True
                    cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
                    Try
                        dgv("CHARVALUE", dgv.Rows.Count - 1) = cbx
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                    End Try

                End If

                'darken cells
                Call SASReadOnly()

                'fill values
                Call CallFillSASValues()


            Catch ex As Exception
                var1 = ex.Message
            End Try



        Catch ex As Exception

            var1 = ex.Message


        End Try


    End Sub

    Sub SASReadOnly()

        Dim var1

        Try

            Dim dgv As DataGridView = Me.dgvSAS

            Dim dv As DataView = dgv.DataSource

            Dim dgv2 As DataGridView = Me.dgvReportTables

            Dim id As Int64
            Dim intRow As Integer
            Dim strF1 As String
            Dim strF2 As String
            Dim str1 As String
            Dim str2 As String
            Dim idCT As Int64
            Dim Count1 As Short

            Dim int1 As Short
            Dim int2 As Short
            Dim strC As String

            str1 = "QC LLOQ, QC Low, QC Mid (QC Mid-1, QC Mid-2), QC High, QC Diln"
            str2 = "Example Exclude: LLOQ OR Dil"

            intRow = dgv2.CurrentRow.Index

            id = dgv2("ID_TBLREPORTTABLE", intRow).Value
            idCT = dgv2("ID_TBLCONFIGREPORTTABLES", intRow).Value


            strF2 = dv.RowFilter
            Dim boolG As Boolean
            For Count1 = 0 To dgv.Rows.Count - 1
                boolG = False
                strC = dgv("CHARCOLUMNNAME", Count1).Value
                If StrComp(strC, "BOOLUSESTDCOLLABELS", CompareMethod.Text) = 0 Then
                    boolG = True
                Else
                    boolG = ReturnGrey(strC, False)
                End If

                If boolG Then

                    dgv.Item("CHARNOT", Count1).Style.BackColor = SystemColors.ControlDarkDark
                    dgv.Item("CHARNOT", Count1).Style.SelectionBackColor = SystemColors.ControlDarkDark
                    dgv.Item("CHARNOT", Count1).ReadOnly = True

                    Select Case strC
                        Case "BOOLUSESTDCOLLABELS"

                        Case "CHARCALSTD"

                        Case Else
                            dgv.Item("CHAREXAMPLE", Count1).Style.BackColor = SystemColors.ControlDarkDark
                            dgv.Item("CHAREXAMPLE", Count1).Style.SelectionBackColor = SystemColors.ControlDarkDark
                            dgv.Item("CHAREXAMPLE", Count1).ReadOnly = True

                    End Select

                Else

                    Select Case strC

                        Case "BOOLUSESTDCOLLABELS"
                            dgv.Item("CHAREXAMPLE", Count1).Value = str1
                        Case "CHARCALSTD"

                        Case Else
                            dgv.Item("CHAREXAMPLE", Count1).Value = str2

                    End Select

                    Select Case idCT

                        Case Else

                    End Select
                End If


            Next

            GoTo end1

            Try
                int1 = InStr(1, strF2, "BOOLUSESTDCOLLABELS", CompareMethod.Text)

                'darken cells
                If int1 > 0 Then
                    dgv.Item("CHARNOT", 0).Style.BackColor = SystemColors.ControlDarkDark
                    dgv.Item("CHARNOT", 0).Style.SelectionBackColor = SystemColors.ControlDarkDark
                    dgv.Item("CHARNOT", 0).ReadOnly = True

                Else
                    'dgv.Item("CHAREXAMPLE", 0).Style.BackColor = SystemColors.ControlDarkDark
                    dgv.Item("CHAREXAMPLE", 0).Value = str2
                End If


                'do all this based on content


                dgv.Item("CHARNOT", dgv.Rows.Count - 1).Style.BackColor = SystemColors.ControlDarkDark
                dgv.Item("CHARNOT", dgv.Rows.Count - 1).Style.SelectionBackColor = SystemColors.ControlDarkDark
                dgv.Item("CHARNOT", dgv.Rows.Count - 1).ReadOnly = True

                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 1).Style.BackColor = SystemColors.ControlDarkDark
                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 1).Style.SelectionBackColor = SystemColors.ControlDarkDark
                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 1).ReadOnly = True

                dgv.Item("CHARNOT", dgv.Rows.Count - 4).Style.BackColor = SystemColors.ControlDarkDark
                dgv.Item("CHARNOT", dgv.Rows.Count - 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                dgv.Item("CHARNOT", dgv.Rows.Count - 4).ReadOnly = True

                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 4).Style.BackColor = SystemColors.ControlDarkDark
                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 4).ReadOnly = True

                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 2).Value = str2
                dgv.Item("CHAREXAMPLE", dgv.Rows.Count - 3).Value = str2

                For Count1 = 1 To dgv.Rows.Count - 5
                    dgv.Item("CHAREXAMPLE", Count1).Value = str2
                Next

                'additional
                Select Case idCT
                    Case 12 'Diln QCs, Diln Factor

                        dgv.Item("CHARNOT", 2).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 2).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 2).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 2).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 2).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 2).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 2).Value = ""
                        dgv.Refresh()

                    Case 22 'Stock Soln Stability

                        dgv.Item("CHARNOT", 2).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 2).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 2).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 2).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 2).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 2).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 2).Value = ""
                        dgv.Refresh()

                    Case 29 'Long Term Stability

                        dgv.Item("CHARNOT", 3).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 3).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 3).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 3).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 3).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 3).ReadOnly = True

                        dgv.Item("CHARNOT", 4).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 4).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 4).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 4).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 3).Value = ""
                        dgv.Item("CHAREXAMPLE", 4).Value = ""
                        dgv.Refresh()

                    Case 31 'Add Hoc Stability

                        dgv.Item("CHARNOT", 2).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 2).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 2).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 2).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 2).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 2).ReadOnly = True

                        dgv.Item("CHARNOT", 3).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 3).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 3).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 3).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 3).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 3).ReadOnly = True

                        'dgv.Item("CHARNOT", 4).Style.BackColor = SystemColors.ControlDarkDark
                        'dgv.Item("CHARNOT", 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        'dgv.Item("CHARNOT", 4).ReadOnly = True

                        'dgv.Item("CHAREXAMPLE", 4).Style.BackColor = SystemColors.ControlDarkDark
                        'dgv.Item("CHAREXAMPLE", 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        'dgv.Item("CHAREXAMPLE", 4).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 2).Value = ""
                        dgv.Item("CHAREXAMPLE", 3).Value = ""
                        'dgv.Item("CHAREXAMPLE", 4).Value = ""
                        dgv.Refresh()

                    Case 32 ' Ad Hoc Stability Comparison

                        dgv.Item("CHARNOT", 3).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 3).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 3).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 3).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 3).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 3).ReadOnly = True

                        dgv.Item("CHARNOT", 4).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHARNOT", 4).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 4).Style.BackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 4).Style.SelectionBackColor = SystemColors.ControlDarkDark
                        dgv.Item("CHAREXAMPLE", 4).ReadOnly = True

                        dgv.Item("CHAREXAMPLE", 3).Value = ""
                        dgv.Item("CHAREXAMPLE", 4).Value = ""
                        dgv.Refresh()

                        'also need to change Run Identifier to Required
                        str1 = "Run Identifier 1 (Required)"
                        dgv.Item("CHARLABEL", 3).Value = str1
                        str1 = "Run Identifier 2 (Required)"
                        dgv.Item("CHARLABEL", 4).Value = str1

                        '20190305 LEE
                        str1 = "Run Identifier 3 (Optional)"
                        dgv.Item("CHARLABEL", 4).Value = str1
                        str1 = "Run Identifier 4 (Optional)"
                        dgv.Item("CHARLABEL", 4).Value = str1


                    Case 35 'Carryover


                End Select

            Catch ex As Exception
                var1 = ex.Message
            End Try

end1:

        Catch ex As Exception

            var1 = ex.Message

        End Try


    End Sub


    Function ReturnSASRows(intTableID As Int64) As String

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String

        Dim bool1 As Boolean

        Dim Count1 As Short

        str1 = "CHARCOLUMNNAME = 'BOOLUSESTDCOLLABELS'"
        'str2 = "CHARCOLUMNNAME = 'CHARSAMPLETYPE' OR CHARCOLUMNNAME = 'CHARRUNDESCR1' OR CHARCOLUMNNAME = 'CHARRUNDESCR2' OR CHARCOLUMNNAME = 'BOOLACCEPTEDONLY'"
        str2 = "CHARCOLUMNNAME = 'CHARSAMPLETYPE' OR CHARCOLUMNNAME = 'CHARRUNDESCR1' OR CHARCOLUMNNAME = 'BOOLACCEPTEDONLY'"

        Select Case intTableID

            Case 2 'Regr Constant table
                str3 = "CHARCOLUMNNAME = 'CHARCALSTD'"
                'str4 = str1 & " OR " & str2 & " OR " & str3
                str4 = str2 & " OR " & str3

            Case 3 'Summary of Back-Calculated Calibration Std Conc
                str3 = "CHARCOLUMNNAME = 'CHARCALSTD'"
                'str4 = str1 & " OR " & str2 & " OR " & str3
                str4 = str2 & " OR " & str3
            Case 4 'Summary of Interpolated QC Std Conc
                str3 = "CHARCOLUMNNAME = 'CHARNONCOREQC'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 11 'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
                str3 = "CHARCOLUMNNAME = 'CHARCOREQC'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 12 'Summary of Interpolated Dilution QC Concentrations
                str3 = "CHARCOLUMNNAME = 'CHARDILN' OR CHARCOLUMNNAME = 'CHARDILNFACTOR'"
                str4 = str1 & " OR " & str2 & " OR " & str3
                'str4 = str2 & " OR " & str3
            Case 13 'Summary of Combined Recovery
                str3 = "CHARCOLUMNNAME = 'CHARRECRS' OR CHARCOLUMNNAME = 'CHARRECQC'"
                str4 = str2 & " OR " & str3
            Case 14 'Summary of True Recovery
                str3 = "CHARCOLUMNNAME = 'CHARRECPES' OR CHARCOLUMNNAME = 'CHARRECQC'"
                str4 = str2 & " OR " & str3
            Case 15 'Summary of Suppression/Enhancement
                str3 = "CHARCOLUMNNAME = 'CHARRECRS' OR CHARCOLUMNNAME = 'CHARRECPES'"
                str4 = str2 & " OR " & str3
            Case 17 'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
                str3 = "CHARCOLUMNNAME = 'CHARLOT1'"
                For Count1 = 2 To 10
                    str3 = str3 & " OR CHARCOLUMNNAME = 'CHARLOT" & Count1 & "'"
                Next
                str4 = str1 & " OR " & str2 & " OR " & str3

            Case 34 'Selectivity in Individual Lots Table v1
                str3 = "CHARCOLUMNNAME = 'CHARLOT1'"
                For Count1 = 2 To 10
                    str3 = str3 & " OR CHARCOLUMNNAME = 'CHARLOT" & Count1 & "'"
                Next
                '20181216 LEE:
                For Count1 = 1 To 10
                    str3 = str3 & " OR CHARCOLUMNNAME = 'CHARLOTWOIS" & Count1 & "'"
                Next
                'str4 = str1 & " OR " & str2 & " OR " & str3
                str4 = str2 & " OR " & str3 & " OR CHARCOLUMNNAME = 'CHARCALSTD'"
                'str4 = str4 & " OR CHARCOLUMNNAME = 'CHARCALSTD'"
            Case 18 'Summary of [Period Temp] Stability in Matrix
                str3 = "CHARCOLUMNNAME = 'CHARNONCOREQC'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 19 'Summary of Freeze/Thaw [#Cycles] Stability in Matrix
                str3 = "CHARCOLUMNNAME = 'CHARNONCOREQC'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 21 '[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
                str3 = "CHARCOLUMNNAME = 'CHARNONCOREQC'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 22 '[Period Temp] Stock Solution Stability Assessment
                str3 = "CHARCOLUMNNAME = 'CHAROLD' OR CHARCOLUMNNAME = 'CHARNEW' OR CHARCOLUMNNAME = 'CHARSTOCKSOLNCONC'"
                'str4 = str1 & " OR " & str2 & " OR " & str3
                str4 = str2 & " OR " & str3
            Case 23 '[Period Temp] Spiking Solution Stability Assessment
                str3 = "CHARCOLUMNNAME = 'CHAROLD' OR CHARCOLUMNNAME = 'CHARNEW'"
                'str4 = str1 & " OR " & str2 & " OR " & str3
                str4 = str2 & " OR " & str3
            Case 29 '[Period Temp] Long-Term QC Std Storage Stability
                str3 = "CHARCOLUMNNAME = 'CHAROLD' OR CHARCOLUMNNAME = 'CHARNEW' OR CHARCOLUMNNAME = 'CHARRUNIDENTIFIER1' OR CHARCOLUMNNAME = 'CHARRUNIDENTIFIER2'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 31 'Ad Hoc QC Stability Table
                str3 = "CHARCOLUMNNAME = 'CHARNONCOREQC'"
                str3 = "CHARCOLUMNNAME = 'CHARNONCOREQC'  OR CHARCOLUMNNAME = 'CHARRUNIDENTIFIER1'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 32 'Ad Hoc QC Stability Comparison Table
                ' str3 = "CHARCOLUMNNAME = 'CHAROLD' OR CHARCOLUMNNAME = 'CHARNEW'"
                str3 = "CHARCOLUMNNAME = 'CHAROLD' OR CHARCOLUMNNAME = 'CHARNEW' OR CHARCOLUMNNAME = 'CHARNEW2' OR CHARCOLUMNNAME = 'CHARNEW3'  OR CHARCOLUMNNAME = 'CHARRUNIDENTIFIER1' OR CHARCOLUMNNAME = 'CHARRUNIDENTIFIER2' OR CHARCOLUMNNAME = 'CHARRUNIDENTIFIER3' OR CHARCOLUMNNAME = 'CHARRUNIDENTIFIER4'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case 35 'Carryover
                str3 = "CHARCOLUMNNAME = 'CHARLLOQ' OR CHARCOLUMNNAME = 'CHARULOQ' OR CHARCOLUMNNAME = 'CHARBLANK'"
                str4 = str1 & " OR " & str2 & " OR " & str3
            Case Else
                str4 = ""

        End Select

        ReturnSASRows = str4

    End Function

    Sub CallFillSASValues()

        Dim dgvS As DataGridView = Me.dgvASP
        Dim dgvD As DataGridView = Me.dgvSAS

        Dim boolSE As Boolean = False

        Dim Count1 As Short
        Dim strCol As String
        Dim var1, var2, var3, var4, var5
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        Dim intLot As Short = 0
        Dim intLotWO As Short = 0

        Dim idC As Int32 = idCGet()

        For Count1 = 0 To dgvD.Rows.Count - 1

            Dim strNot As String = ""

            strCol = dgvD("CHARCOLUMNNAME", Count1).Value

            'get value
            var3 = dgvS(strCol, 0).Value

            int1 = InStr(1, strCol, "CHARLOT", CompareMethod.Text)
            If int1 > 0 Then
                Select Case idC
                    Case 17
                        'str1 = "Auto Assign Samples Label: Lot 10"
                        If Me.chkBOOLDOINDREC.Checked Then

                            If intLot = 0 Then
                                str2 = "Solvent"
                            Else
                                str2 = "Lot " & intLot
                            End If
                            'str1 = "Auto Assign Samples Label: " & str2
                            str1 = "Sample Name text: " & str2
                            dgvD("CHARLABEL", Count1).Value = str1
                        Else
                            'str1 = "Auto Assign Samples Label: Lot " & intLot + 1
                            str1 = "Sample Name text: Lot" & intLot + 1
                            dgvD("CHARLABEL", Count1).Value = str1
                        End If
                        intLot = intLot + 1

                    Case 34 'Selectivity Lots

                        int2 = InStr(1, strCol, "CHARLOTWOIS", CompareMethod.Text)

                        If int2 > 0 Then
                            intLotWO = intLotWO + 1
                            'str1 = "Auto Assign Samples Label WithOut IS Lot " & intLotWO
                            str1 = "Sample Name text: WithOut IS Lot " & intLotWO
                        Else
                            intLot = intLot + 1
                            'str1 = "Auto Assign Samples Label With IS: Lot " & intLot
                            str1 = "Sample Name text: With IS Lot " & intLot
                        End If

                        dgvD("CHARLABEL", Count1).Value = str1

                End Select

            End If

            'look for more special


            Select Case idC

                Case 35
                    int1 = InStr(1, strCol, "CHARULOQ", CompareMethod.Text)
                    If int1 > 0 Then
                        str1 = "Sample Name text (Optional):"
                        str2 = "ULOQ"
                        str3 = dgvD("CHARLABEL", Count1).Value
                        If InStr(1, str3, str2, CompareMethod.Text) > 0 Then
                            str1 = "Sample Name text (Optional): ULOQ"
                            dgvD("CHARLABEL", Count1).Value = str1
                        End If
                    End If
            End Select


            'strip Not portion
            int1 = InStr(1, var3.ToString, "(", CompareMethod.Text)
            If int1 = 0 Then
                var1 = var3
            Else
                var1 = Mid(var3, 1, int1 - 1)
                strNot = Mid(var3, int1 + 1, Len(var3) - int1 - 1)
            End If

            Select Case strCol

                Case "BOOLUSESTDCOLLABELS"
                    var2 = NZ(var1, 0)
                    boolSE = True
                    If var2 = 0 Then
                        dgvD("CHARVALUE", Count1).Value = "FALSE"
                    Else
                        dgvD("CHARVALUE", Count1).Value = "TRUE"
                    End If

                Case "BOOLACCEPTEDONLY"
                    var2 = NZ(var1, 0)
                    'boolSE = True
                    If var2 = 0 Then
                        dgvD("CHARVALUE", Count1).Value = "FALSE"
                    Else
                        dgvD("CHARVALUE", Count1).Value = "TRUE"
                    End If

                Case Else
                    dgvD("CHARVALUE", Count1).Value = var1
                    If Len(strNot) = 0 Then
                    Else
                        dgvD("CHARNOT", Count1).Value = strNot
                    End If

            End Select

        Next

        dgvD.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        dgvD.AutoResizeColumns()

        dgvD.Columns("CHARVALUE").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgvD.Columns("CHARNOT").AutoSizeMode = DataGridViewAutoSizeColumnMode.None



        'If boolSE Then
        '    dgvD.Columns("CHAREXAMPLE").Visible = True
        'Else
        '    dgvD.Columns("CHAREXAMPLE").Visible = False
        'End If



    End Sub


    Private Sub txtTitle_TextChanged_1(sender As Object, e As EventArgs) Handles txtTitle.TextChanged

    End Sub

    Private Sub txtTitle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTitle.KeyPress
        If e.KeyChar = Convert.ToChar(1) Then
            DirectCast(sender, TextBox).SelectAll()
            e.Handled = True
        End If
    End Sub

    Private Sub rbDontShowRejected_CheckedChanged(sender As Object, e As EventArgs) Handles rbDontShowRejected.CheckedChanged

    End Sub

    Private Sub cmdSymbol_Click(sender As Object, e As EventArgs) Handles cmdSymbol.Click

        Dim frm As New frmShowSymbol
        'place form to right of form

        Dim a, b, c, d

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height


        a = Me.ClientSize.Width
        a = Me.cmdSymbol.Left
        b = a ' a - frm.Width

        frm.Left = b

        c = Me.cmdSymbol.Top + Me.cmdSymbol.Height + tbh + 2
        frm.Top = c

        frm.Show()


    End Sub

    Private Sub chkBOOLDOINDREC_CheckedChanged(sender As Object, e As EventArgs) Handles chkBOOLDOINDREC.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkBOOLDOINDREC, "BOOLDOINDREC")
            Call UpdateMF01()
            Call UpdateME01()
        End If

    End Sub

    Sub UpdateME01()

        Dim idC As Int32
        Dim t1, l1
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim l

        idC = idCGet()

        str1 = "Matrix Factor"
        str2 = "Show table as Matrix Factor"
        str3 = "Include Matrix Factor columns for Analyte and Int Std"
        str4 = "Include IntStd-Normalized Matrix Factor column"

        l = 37

        Select Case idC
            Case 17

                l1 = Me.gbResultsChoice.Left + Me.gbResultsChoice.Width + 5
                t1 = Me.gbResultsChoice.Top

                l = 15

                Me.gbMatrixFactor.Top = t1
                Me.gbMatrixFactor.Left = l1
                'Me.gbMatrixFactor.Height = Me.gbResultsChoice.Height

                If Me.chkBOOLDOINDREC.Checked Then

                    Me.rbConc.Visible = False
                    If Me.rbConc.Checked Then
                        Me.rbUsePeakArea.Checked = True
                        Me.rbConc.Checked = False
                    End If
                    Me.chkBOOLISCOMBINELEVELS.Visible = False
                    Me.chkBOOLISCOMBINELEVELS.Checked = False

                    Me.chkInclIntStdNMF.Visible = False

                    str1 = "Matrix Effect and Matrix Factor"
                    str2 = "Include Overall Mean MF statistics section"
                    str3 = "Calculate Mean MF as average of individual MF calculations"
                    str4 = "Refer to Matrix Factor as Matrix Effect"

                    Me.gbMatrixFactor.Visible = True
                    Me.gbMatrixFactor.Text = str1
                    Me.chkMFTable.Text = str2
                    Me.chkInclMFCols.Text = str3
                    Me.chkInclIntStdNMF.Text = str4
                    Me.chkInclIntStdNMF.Visible = True
                    Me.panMF1.Left = l

                    Me.panMF1.Enabled = True

                Else
                    Me.gbMatrixFactor.Visible = False
                    Me.rbConc.Visible = True

                End If

            Case Else
                Me.chkInclIntStdNMF.Visible = True
                Me.gbMatrixFactor.Text = str1
                Me.chkMFTable.Text = str2
                Me.chkInclMFCols.Text = str3
                Me.chkInclIntStdNMF.Text = str4
                Me.panMF1.Left = l
        End Select

    End Sub

    Function idCGet() As Int32

        Dim dgv As DataGridView
        Dim intRow As Int16
        Dim idC As Int32

        dgv = Me.dgvReportTables

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        Try
            idC = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        Catch ex As Exception
            idC = 0
        End Try

        idCGet = idC

    End Function

    Sub UpdateMF01()

        Dim strM As String

        Dim idC As Int32 = idCGet()

        Select Case idC
            Case 17
                'Me.lblMF01.Visible = False
            Case Else
                If Me.chkInclMFCols.Checked Then
                    If Me.chkBOOLDOINDREC.Checked Then
                        strM = "IntStd-Normalized Matrix Factor (MF) calculated as" & ChrW(10) & "MF(Analyte)/MF(IntStd)"
                    Else
                        strM = "IntStd-Normalized Matrix Factor (MF) calculated as" & ChrW(10) & "(Mean MF(Analyte))/(Mean MF(IntStd))"
                    End If
                Else
                    If Me.chkBOOLDOINDREC.Checked Then
                        strM = "IntStd-Normalized Matrix Factor (MF) calculated as" & ChrW(10) & "(Peak Area Ratio PES)/(Peak Area Ratio RS)"
                    Else
                        strM = "IntStd-Normalized Matrix Factor (MF) calculated as" & ChrW(10) & "(Mean Peak Area Ratio PES)/(Mean Peak Area Ratio RS)"
                    End If
                End If

                Me.lblMF01.Text = strM
                Me.lblMF01.Visible = True

        End Select

        ''Legend
        'Select Case Count1
        '    Case 1
        '        strFirst = "Matrix Factor ="
        '        strNum = "(Mean Peak Area Ratio Post Extraction Spiking Solution)"
        '        strDenom = "(Mean Peak Area Ratio Recovery Solution)"
        '    Case 2
        '        strFirst = NZ(rows(0).Item("CHARTITLELEG"), "")
        '        strNum = NZ(rows(0).Item("CHARNUMLEG"), "")
        '        strDenom = NZ(rows(0).Item("CHARDENLEG"), "")
        'End Select


    End Sub

    Private Sub dgvSAS_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvSAS.CellValidating

        'update dgvASP

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgvS As DataGridView = Me.dgvSAS
        Dim dgvD As DataGridView = Me.dgvASP

        Dim intRow As Int32
        Dim intCol As Short
        Dim strCol As String
        Dim strM As String
        Dim str1 As String
        Dim boolUSCL As Boolean

        intRow = e.RowIndex
        intCol = e.ColumnIndex

        Dim boolIsNOTCol As Boolean = False
        Dim strColSAS As String = dgvS.Columns(intCol).Name

        If StrComp(strColSAS, "CHARNOT", CompareMethod.Text) = 0 Then
            boolIsNOTCol = True
        End If

        If StrComp(strColSAS, "CHARVALUE", CompareMethod.Text) = 0 Or boolIsNOTCol Then
            If dgvS.Columns(intCol).ReadOnly Then
                GoTo end1
            End If
        Else
            GoTo end1
        End If


        'get tblAutoAssignSample column
        strCol = dgvS("CHARCOLUMNNAME", intRow).Value

        Select Case strCol
            Case "BOOLUSESTDCOLLABELS", "BOOLACCEPTEDONLY"
                boolUSCL = True
            Case Else
                boolUSCL = False
        End Select

        Dim vValO
        If boolIsNOTCol Then
            vValO = dgvS("CHARNOT", intRow).Value
        Else
            vValO = dgvS("CHARVALUE", intRow).Value
        End If


        'get value
        Dim vVal, vNot
        If boolIsNOTCol Then
            vVal = dgvS("CHARVALUE", intRow).Value
            vNot = e.FormattedValue
        Else
            vVal = e.FormattedValue
            vNot = dgvS("CHARNOT", intRow).Value
        End If


        'Dim intSpC As Short = HasSpecialCharacters(vVal.ToString)
        Dim intSpC As Short = HasSpecialCharacters(e.FormattedValue.ToString)


        If intSpC = 1 Then
            strM = "The entry '" & e.FormattedValue & "' contains " & intSpC & " invalid character."
            MsgBox(strM, vbInformation, "Invalid action...")
            e.Cancel = True
            dgvS(strColSAS, intRow).Value = vValO
            dgvS.Refresh()
        Else
            'put value
            'value must be 20 characters or less if CHARVALUE
            If StrComp(strColSAS, "CHARVALUE", CompareMethod.Text) = 0 Then
                'This is unnecessary
                'If Len(vVal) > 20 Then
                '    strM = "The entry '" & vVal.ToString & "' contains " & Len(vVal) & " characters; maximum allowed is 20."
                '    MsgBox(strM, vbInformation, "Invalid action...")
                '    e.Cancel = True
                '    GoTo end1

                'End If

            End If

            Try

                Try
                    'Hmmm. Database doesn't seem to be updating if data in grid is changed.
                    'must updated data in dv
                    Dim dv As DataView = dgvASP.DataSource
                    dv(0).BeginEdit()
                    If boolUSCL And boolIsNOTCol = False Then
                        If StrComp(vVal, "FALSE", CompareMethod.Text) = 0 Then
                            dv(0).Item(strCol) = 0
                        Else
                            dv(0).Item(strCol) = -1
                        End If
                    Else

                        ''Legend
                        'If boolIsNOTCol Then
                        '    vVal = dgvS("CHARVALUE", intRow).Value
                        '    vNot = e.FormattedValue
                        'Else
                        '    vVal = e.FormattedValue
                        '    vNot = dgvS("CHARNOT", intRow).Value
                        'End If

                        Dim vF
                        Dim strNot As String = ""

                        'evaluate Exclusions
                        If dgvS("CHARNOT", intRow).ReadOnly Then
                            vF = vVal
                        Else
                            If Len(NZ(vNot, "")) = 0 Then
                                strNot = "()"
                            Else
                                strNot = "(" & vNot & ")"
                            End If
                        End If
                        vF = vVal & strNot

                        dv(0).Item(strCol) = vF ' vVal

                    End If

                    dv(0).EndEdit()

                Catch ex As Exception
                    Dim v1
                    v1 = ex.Message
                End Try

            Catch ex As Exception

            End Try

        End If


end1:


    End Sub

    Private Sub tabRTC_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabRTC.SelectedIndexChanged

        'pesky
        Dim dgv As DataGridView = Me.dgvSAS
        dgv.AutoResizeRows()

        Call FilterSAS()

    End Sub

    Private Sub cmdShowRunSummary_Click(sender As Object, e As EventArgs) Handles cmdShowRunSummary.Click

        'MsgBox("Under construction", MsgBoxStyle.Information, "Under construction...")

        Dim frm As New frmAnalyticalRunSummary

        frm.Show(Me)

    End Sub

    Private Sub cmdAnalRuns_Click(sender As Object, e As EventArgs) Handles cmdAnalRuns.Click

        Call OpenAssignedSamples(True)

    End Sub

    Private Sub chkBOOLREASSAYREASLETTERS_CheckedChanged(sender As Object, e As EventArgs) Handles chkBOOLREASSAYREASLETTERS.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkBOOLREASSAYREASLETTERS, "BOOLREASSAYREASLETTERS")
        End If

    End Sub

    Private Sub chkBOOLISCOMBINELEVELS_CheckedChanged(sender As Object, e As EventArgs) Handles chkBOOLISCOMBINELEVELS.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkBOOLISCOMBINELEVELS, "BOOLISCOMBINELEVELS")
        End If

    End Sub

    Private Sub cmdTest_Click(sender As Object, e As EventArgs) Handles cmdTest.Click

        MsgBox(INTQCLEVELGROUP.ToString)

    End Sub

    Private Sub chkIncludeIS_Single_CheckedChanged(sender As Object, e As EventArgs) Handles chkIncludeIS_Single.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkIncludeIS_Single, "BOOLINCLUDEISTBL")

        End If

    End Sub

    Private Sub CHARTITLELEG_TextChanged(sender As Object, e As EventArgs) Handles CHARTITLELEG.TextChanged

    End Sub

    Private Sub cmdPasteConditions_Click(sender As Object, e As EventArgs) Handles cmdPasteConditions.Click

        Call PasteConditions()

    End Sub

    Sub PasteConditions()

        Dim strM As String
        Dim strM1 As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim boolE As Boolean = False
        Dim strE As String = ""
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim intCR As Short
        Dim intTabs As Short
        Dim intNumTabs As Short = 13 ' 12. 20190305 LEE: Added column for run identifier for Ad Hoc Stab Comparison
        Dim intL As Int64
        Dim strP As String = ""
        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim Count3 As Int16
        Dim var1, var2, var3

        Dim idCRT As Int64 'id_tblConfigReporttables
        Dim idRT As Int64 'ID_TBLREPORTTABLE
        Dim idRT1 As Int64


        str1 = "Paste Rules: The clipboard should contain " & intNumTabs & " tab-delimeted columns (no column headers)." & ChrW(10) & ChrW(10)
        str1 = str1 & "The column contents should be:" & ChrW(10)
        int1 = 0
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".   Table Order"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".   Placeholder (X)"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".   Include (X)"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  StudyDoc FC ID (e.g. BTSTAB1) REQUIRED"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  Table Title"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  # of Cycles (for Freeze/Thaw)"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  Period Length (e.g. '4' or 'four')"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  Period Units (e.g 'hours')"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  Temperature Description (e.g. -80 °C)"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  Stability Description (combination of Length, Units and Temp)"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ".  Sample Name Text Fragment (e.g. FT3C80)"
        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ". Analytical Run Description Text Fragment (e.g. Reinj)"

        int1 = int1 + 1
        str1 = str1 & ChrW(10) & strP.PadRight(5) & Format(int1, "00") & ". Run Identifier for %Difference tables (e.g. '0 Minutes')"

        str1 = str1 & ChrW(10) & ChrW(10) & "Note that blank entries will be ignored."
        str1 = str1 & ChrW(10) & ChrW(10) & "The paste contents will be validated after clicking the OK button."
        Dim frm As New frmPasteConditions

        frm.lbl1.Text = str1
        frm.ShowDialog()
        Dim boolCancel As Boolean = frm.boolCancel
        frm.Dispose()

        If boolCancel Then
            GoTo end2
        End If

        Try

            Dim strC As String = My.Computer.Clipboard.GetText()

            intL = Len(strC)
            'validate contents
            If intL = 0 Then
                boolE = True
                strM1 = "The clipboard is empty or does not contain text."
                GoTo end1
            End If

            'look for number of tabs: should be 7
            intTabs = 0

            Dim intL1 As Int16
            Dim intL2 As Int16
            Dim intL3 As Int16
            Dim intUB1 As Short
            Dim intUB2 As Short
            Dim arrC(8, 1000)
            'Dim arrTabs(1000)


            Dim boolCTime As Boolean = Me.chkCONVERTTIME.Checked
            Dim boolCTemp As Boolean = Me.chkCONVERTTEMP.Checked

            Dim strCR As String = ChrW(13) & ChrW(10)

            Dim arrTabs(1000)
            '20190305 LEE:
            'Use .split instead
            Dim v() As String = Split(strC, ChrW(13) & ChrW(10)) 'split returns a 0-based array
            'This will return an extra blank line if pasted from Excel
            'also need to convert to 1-based array
            intCR = UBound(v)
            ReDim arrTabs(intCR)
            For Count2 = 0 To v.Length - 2
                var1 = v(Count2)
                arrTabs(Count2 + 1) = var1
            Next

            ReDim arrC(8, intCR)

            'now find number of columns
            'use first arrTabs entry
            'remember that last linefeed/carriage return has been replaced with a tab character
            Dim strT As String

            '20190305 LEE: Use split
            'check for correct columns
            var1 = arrTabs(Count2)
            Dim w() As String = Split(CStr(var1), ChrW(9)) 'split returns a 0-based array
            intTabs = UBound(w) + 1

            If intTabs <> intNumTabs Then
                strM1 = "The paste text contains " & intTabs & " cells."
                strM1 = strM1 & ChrW(10) & "Rules require " & intNumTabs & " cells."
                boolE = True
                GoTo end1
            End If

            'now parse contents to table
            Dim dtbl As New DataTable 'all columns will be type = string
            For Count1 = 1 To intNumTabs

                Select Case Count1
                    Case 1
                        str1 = "Order"
                    Case 2
                        str1 = "PlaceHolder"

                    Case 3
                        str1 = "Include"
                    Case 4
                        str1 = "FCID"
                    Case 5
                        str1 = "TableTitle"
                    Case 6
                        str1 = "NumCycles"
                    Case 7
                        str1 = "PL" 'Period Length
                    Case 8
                        str1 = "PU" 'Period Units
                    Case 9
                        str1 = "TD" 'Temperature Description
                    Case 10
                        str1 = "SD" 'Stability Descr
                    Case 11
                        str1 = "SN" 'Sample Name Text Fragment
                    Case 12
                        str1 = "AR" 'Analytical Run Description Text Fragment

                    Case 13 '20190305 LEE:
                        str1 = "RI" 'Run Identifier
                End Select

                Dim col1 As New DataColumn
                col1.ColumnName = str1
                col1.AllowDBNull = True
                dtbl.Columns.Add(col1)

            Next Count1

            'now populate dtbl
            For Count1 = 1 To intCR

                str1 = arrTabs(Count1)

                '20190305 LEE: use split
                Dim x() As String = Split(CStr(str1), ChrW(9)) 'split returns a 0-based array
                intNumTabs = UBound(x) + 1

                Dim nr As DataRow = dtbl.NewRow
                nr.BeginEdit()
                Try
                    For Count2 = 1 To intNumTabs

                        var1 = x(Count2 - 1)

                        nr(Count2 - 1) = var1
                        int2 = int1 + 1
                    Next Count2
                Catch ex As Exception
                    var2 = ex.Message
                End Try

                nr.EndEdit()
                dtbl.Rows.Add(nr)

            Next Count1

            Dim strFCID As String
            Dim strF As String
            Dim strF1 As String
            Dim strF2 As String
            Dim strF3 As String

            Dim strFCID2 As String

        
            'Dim dgvRT As DataGridView = Me.dgvReportTables
            'Dim dvRT As DataView = dgvRT.DataSource

            ''must accept table first
            'tblReportTables.AcceptChanges()

            'can't use dvRT count because changing BOOLINCLUDE increase or decreases number of rows
            'must use straight tblReportTable
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            Dim rowsRT() As DataRow = tblReportTables.Select(strF)

            For Count2 = 0 To rowsRT.Length - 1

                strFCID = NZ(rowsRT(Count2).Item("CHARFCID"), "")
                If Len(strFCID) = 0 Then
                    strFCID2 = ""
                Else
                    strFCID2 = Mid(strFCID, 1, 2)
                End If
                idRT = rowsRT(Count2).Item("ID_TBLREPORTTABLE")
                idCRT = rowsRT(Count2).Item("ID_TBLCONFIGREPORTTABLES") 'do things based table type


                If idCRT = 32 Then
                    idCRT = idCRT 'debug
                End If

                If Len(strFCID) = 0 Then
                Else
                    strF1 = "FCID = '" & strFCID & "'"
                    Dim rows() As DataRow = dtbl.Select(strF1)
                    If rows.Length = 0 Then
                    Else

                        rowsRT(Count2).BeginEdit()

                        '20181218 LEE:
                        'Do not apply order
                        'If there are Un-included rows in paste source, order get's messed up
                        ''
                        'var1 = NZ(rows(0).Item("Order"), "")
                        'If IsNumeric(var1) Then
                        '    rowsRT(Count2).Item("INTORDER") = CInt(var1)
                        'Else
                        '    var2 = NZ(rowsRT(Count2).Item("INTORDER"), "")
                        '    If IsNumeric(var2) Then
                        '    Else
                        '        rowsRT(Count2).Item("INTORDER") = 100
                        '    End If
                        'End If


                        var1 = NZ(rows(0).Item("Placeholder"), "")
                        If Len(var1) = 0 Then
                            rowsRT(Count2).Item("BOOLPLACEHOLDER") = 0
                        Else
                            rowsRT(Count2).Item("BOOLPLACEHOLDER") = -1
                        End If


                        var1 = NZ(rows(0).Item("Include"), "")
                        If Len(var1) = 0 Then
                            rowsRT(Count2).Item("BOOLINCLUDE") = 0
                        Else
                            rowsRT(Count2).Item("BOOLINCLUDE") = -1
                        End If

                        var1 = NZ(rows(0).Item("TableTitle"), "")
                        If Len(var1) = 0 Then
                        Else
                            rowsRT(Count2).Item("CHARHEADINGTEXT") = var1
                        End If

                        '20181219 LEE
                        Select Case idCRT
                            Case 32
                                var1 = rows(rows.Length - 1).Item("SD")
                            Case Else
                                var1 = rows(0).Item("SD")
                        End Select
                        If Len(var1) = 0 Then
                        Else
                            rowsRT(Count2).Item("CHARSTABILITYPERIOD") = var1
                        End If


                        strF2 = "ID_TBLREPORTTABLE = " & idRT
                        Dim rowsTP() As DataRow = tblTableProperties.Select(strF2)
                        If rowsTP.Length = 0 Or StrComp(strFCID2, "PH", CompareMethod.Text) = 0 Then 'ignore placeholder tables
                        Else

                            rowsTP(0).BeginEdit()
                            Select Case idCRT
                                Case 19 'Freeze/thaw
                                    var1 = NZ(rows(0).Item("NumCycles"), "")
                                    If Len(var1) = 0 Or IsNumeric(var1) = False Then
                                    Else
                                        rowsTP(0).Item("INTNUMBEROFCYCLES") = var1
                                    End If
                            End Select



                            '20181219 LEE
                            Select Case idCRT
                                Case 32 'Ad Hoc Stability Comparison
                                    var1 = rows(rows.Length - 1).Item("PL")
                                Case Else
                                    var1 = rows(0).Item("PL")
                            End Select
                            If Len(var1) = 0 Then
                            Else
                                var2 = var1
                                rowsTP(0).Item("CHARTIMEPERIOD") = var2
                            End If

                            '20181219 LEE
                            Select Case idCRT
                                Case 32 'Ad Hoc Stability Comparison
                                    var1 = rows(rows.Length - 1).Item("PU")
                                Case Else
                                    var1 = rows(0).Item("PU")
                            End Select
                            If Len(var1) = 0 Then
                            Else
                                rowsTP(0).Item("CHARTIMEFRAME") = var1
                            End If

                            '20181219 LEE
                            Select Case idCRT
                                Case 32 'Ad Hoc Stability Comparison
                                    var1 = rows(rows.Length - 1).Item("TD")
                                Case Else
                                    var1 = rows(0).Item("TD")
                            End Select
                            If Len(var1) = 0 Then
                            Else
                                var2 = var1
                                rowsTP(0).Item("CHARPERIODTEMP") = var2
                            End If

                            'since this is coming from an Excel file, must set the following to false
                            rowsTP(0).Item("BOOLCONVERTTIME") = 0
                            rowsTP(0).Item("BOOLCONVERTTEMP") = 0

                            rowsTP(0).EndEdit()

                        End If


                        Dim rowsAAS() As DataRow = tblAutoAssignSamples.Select(strF2)
                        If rowsAAS.Length = 0 Then
                        Else

                            rowsAAS(0).BeginEdit()

                            var1 = rows(0).Item("SN")
                            Select Case idCRT
                                Case 35 'carryover, can have two or three rows

                                    Try
                                        If rows.Length = 2 Then
                                            var1 = rows(0).Item("SN")
                                            If Len(var1) = 0 Then
                                            Else
                                                rowsAAS(0).Item("CHARLLOQ") = var1
                                            End If
                                            var1 = rows(1).Item("SN")
                                            If Len(var1) = 0 Then
                                            Else
                                                rowsAAS(0).Item("CHARBLANK") = var1
                                            End If
                                        Else
                                            var1 = rows(0).Item("SN")
                                            If Len(var1) = 0 Then
                                            Else
                                                rowsAAS(0).Item("CHARLLOQ") = var1
                                            End If
                                            var1 = rows(1).Item("SN")
                                            If Len(var1) = 0 Then
                                            Else
                                                rowsAAS(0).Item("CHARULOQ") = var1
                                            End If
                                            var1 = rows(2).Item("SN")
                                            If Len(var1) = 0 Then
                                            Else
                                                rowsAAS(0).Item("CHARBLANK") = var1
                                            End If
                                        End If
                                    Catch ex As Exception
                                        var1 = var1
                                    End Try
                                   
                                Case 13 'combined recovery
                                    var1 = rows(0).Item("SN")
                                    If Len(var1) = 0 Then
                                    Else
                                        rowsAAS(0).Item("CHARRECRS") = var1
                                    End If
                                    var1 = rows(1).Item("SN")
                                    If Len(var1) = 0 Then
                                    Else
                                        rowsAAS(0).Item("CHARRECQC") = var1
                                    End If
                                Case 14 'true recovery
                                    var1 = rows(0).Item("SN")
                                    If Len(var1) = 0 Then
                                    Else
                                        rowsAAS(0).Item("CHARRECPES") = var1
                                    End If
                                    var1 = rows(1).Item("SN")
                                    If Len(var1) = 0 Then
                                    Else
                                        rowsAAS(0).Item("CHARRECQC") = var1
                                    End If
                                Case 15 'suppr/enhn/matrixfactor
                                    var1 = rows(0).Item("SN")
                                    If Len(var1) = 0 Then
                                    Else
                                        rowsAAS(0).Item("CHARRECPES") = var1
                                    End If
                                    var1 = rows(1).Item("SN")
                                    If Len(var1) = 0 Then
                                    Else
                                        rowsAAS(0).Item("CHARRECRS") = var1
                                    End If
                                Case 17 'Unique lots; 
                                    For Count3 = 1 To 10
                                        str1 = "CHARLOT" & Count3
                                        var1 = rows(Count3 - 1).Item("SN")
                                        If Len(var1) = 0 Then
                                        Else
                                            rowsAAS(0).Item(str1) = var1
                                        End If
                                    Next
                                Case 34 'Selectivity, has 21 rows, 20181207 LEE
                                    For Count3 = 1 To 10
                                        str1 = "CHARLOT" & Count3
                                        var1 = rows(Count3 - 1).Item("SN")
                                        If Len(var1) = 0 Then
                                        Else
                                            rowsAAS(0).Item(str1) = var1
                                        End If
                                    Next
                                    For Count3 = 11 To 20
                                        str1 = "CHARLOTWOIS" & Count3 - 10
                                        var1 = rows(Count3 - 1).Item("SN")
                                        If Len(var1) = 0 Then
                                        Else
                                            rowsAAS(0).Item(str1) = var1
                                        End If
                                    Next
                                    'also has LLOQ Std
                                    var1 = rows(20).Item("SN")
                                    rowsAAS(0).Item("CHARCALSTD") = var1
                                Case 11 'Core QCs
                                    rowsAAS(0).Item("CHARCOREQC") = var1

                                Case 32 'Ad Hoc QC Stability Comparison Table
                                    'has CHAROLD
                                    'has CHARNEW
                                    'for now, use only charnew
                                    'rowsAAS(0).Item("CHARNEW") = var1

                                    'has 8 more rows

                                    '20190305 LEE:
                                    For Count3 = 1 To 4
                                        Select Case Count3
                                            Case 1
                                                str1 = "CHAROLD"
                                            Case 2
                                                str1 = "CHARNEW"
                                            Case 3
                                                str1 = "CHARNEW2"
                                            Case 4
                                                str1 = "CHARNEW3"
                                        End Select

                                        Try
                                            var1 = NZ(rows(Count3 - 1).Item("SN"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                Try
                                                    rowsAAS(0).Item(str1) = var1
                                                Catch ex As Exception
                                                    var1 = var1
                                                End Try

                                            End If
                                        Catch ex As Exception
                                            var1 = var1
                                        End Try
                                     

                                    Next Count3

                                    For Count3 = 1 To 4
                                        Select Case Count3
                                            Case 1
                                                str1 = "CHARRUNIDENTIFIER1"
                                            Case 2
                                                str1 = "CHARRUNIDENTIFIER2"
                                            Case 3
                                                str1 = "CHARRUNIDENTIFIER3"
                                            Case 4
                                                str1 = "CHARRUNIDENTIFIER4"
                                        End Select

                                        Try
                                            var1 = NZ(rows(Count3 - 1).Item("RI"), "")
                                            If Len(var1) = 0 Then
                                            Else
                                                rowsAAS(0).Item(str1) = var1
                                                Try
                                                    rowsAAS(0).Item(str1) = var1
                                                Catch ex As Exception
                                                    var1 = var1
                                                End Try
                                            End If
                                        Catch ex As Exception
                                            var1 = var1
                                        End Try
                                      

                                    Next Count3

                                Case 3, 2 'Calibr Std Table, 20190305 LEE: Need to allow Regr Constant table (2)
                                    '"CHARCALSTD"
                                    rowsAAS(0).Item("CHARCALSTD") = var1
                                Case Else
                                    If Len(var1) = 0 Then
                                    Else
                                        rowsAAS(0).Item("CHARNONCOREQC") = var1
                                    End If

                            End Select




                            var1 = NZ(rows(0).Item("AR"), "")
                            If Len(var1) = 0 Then
                            Else
                                rowsAAS(0).Item("CHARRUNDESCR1") = var1
                            End If

                            rowsAAS(0).EndEdit()

                        End If

                        rowsRT(Count2).EndEdit()

                    End If
                End If
            Next Count2

            'dgvRT is already updated because working with dgvRT.datasource dataview
            'need to update Stability Conditions and AutoAssignment with code
            Call FillProperties()
            Call FilterSAS()

end1:

            If boolE Then
                strM = "The 'Paste Conditions' function was not successful:" & ChrW(10) & ChrW(10) & strM1
                str1 = "Problem..."
            Else
                strM = "Paste Table Names, Conditions, and Auto-Assignments completed successfully."
                str1 = "Done..."
            End If
            MsgBox(strM, vbInformation, str1)

        Catch ex As Exception
            strM = "There was a problem retrieving Report Table Title and Stability information."
            strM = strM & ChrW(10) & ChrW(10) & "Clicking OK will clear the clipboard of all contents."
            strM = strM & ChrW(10) & ChrW(10) & "Please return to the Copy source and attempt to copy again."
            strM = strM & ChrW(10) & ChrW(10) & "idCRT: " & idCRT
            strM = strM & ChrW(10) & ChrW(10) & "Error: " & ex.Message
            MsgBox(strM, vbInformation, "Problem...")
            Try
                Clipboard.Clear()
            Catch ex1 As Exception

            End Try
        End Try

end2:

    End Sub

    Sub DoMFTable()

        Dim idC As Int32

        Try
            idC = idCGet()
            Select Case idC
                Case 17
                    GoTo end1
            End Select
        Catch ex As Exception

        End Try


        If Me.gbMatrixFactor.Visible Then

            If Me.chkMFTable.Checked Then
                Me.panMF1.Enabled = True
                Me.gbTableLegend.Visible = True ' False
            Else
                Me.panMF1.Enabled = False
                Me.gbTableLegend.Visible = True
            End If

            Dim str1 As String
            If Me.chkMFTable.Checked Then
                str1 = "Calculate individual" & ChrW(10) & "Matrix Factor values"
                Me.chkIncludeIS.Enabled = False
                Me.chkIncludeIS.Checked = False
                Me.chkCustomLeg.Enabled = False
                Me.chkCustomLeg.Checked = False
                If Me.chkCalcIntStdNMF.Checked Then
                    Me.rbUsePeakArea.Visible = True
                    Me.rbUsePeakAreaRatio.Visible = False
                    Me.rbUsePeakArea.Checked = True
                    Me.rbUsePeakAreaRatio.Checked = False

                Else
                    Me.rbUsePeakArea.Visible = True
                    Me.rbUsePeakAreaRatio.Visible = False
                    Me.rbUsePeakArea.Checked = True
                    Me.rbUsePeakAreaRatio.Checked = False
                End If

            Else
                str1 = "Calculate individual" & ChrW(10) & "Sup/Enh values"
                Me.rbUsePeakArea.Visible = True
                Me.rbUsePeakAreaRatio.Visible = True
                Me.chkIncludeIS.Enabled = True
            End If
            Me.chkBOOLDOINDREC.Text = str1

        Else

            Me.rbUsePeakArea.Visible = True
            Me.rbUsePeakAreaRatio.Visible = True

        End If

        'Call PARChange(True)

end1:

    End Sub

    Private Sub chkMFTable_CheckedChanged(sender As Object, e As EventArgs) Handles chkMFTable.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkMFTable, "BOOLMFTABLE")
            Call DoMFTable()
        End If

    End Sub

    Private Sub chkInclMFCols_CheckedChanged(sender As Object, e As EventArgs) Handles chkInclMFCols.CheckedChanged

        If boolHold Then
        Else
            If CheckMF(Me.chkInclMFCols) Then
                Call UpdateChkData(False, Me.chkInclMFCols, "BOOLINCLMFCOLS")
                Call UpdateMF01()
                Call DolblMF()
            End If
        End If

    End Sub

    Sub DolblMF()

        Dim idC As Int32 = idCGet()
        Dim str1 As String
        Dim str2 As String
        Select Case idC
            Case 17
                If Me.chkBOOLDOINDREC.Checked Then
                    If Me.chkInclMFCols.Checked Then
                        str1 = "Mean Matrix Factor = (Average of Matrix Factor column)"
                    Else
                        str1 = "Mean Matrix Factor = (Average Lot Peak Area)/(Average Solvent Peak Area)"
                    End If

                    If Me.chkInclIntStdNMF.Checked Then

                        If Me.chkInclMFCols.Checked Then
                            str1 = "Mean Matrix Effect = (Average of Matrix Effect column)"
                        Else
                            str1 = "Mean Matrix Effect = (Average Lot Peak Area)/(Average Solvent Peak Area)"
                        End If
                    Else


                        '20181108 LEE: Depracate this
                        'Only used by CRL-WIL, and they have stopped using it
                        'str2 = "Matrix Effect = ((Mean Matrix Factor * 100) - 100)"
                        'str1 = str1 & ChrW(10) & str2
                    End If

                    Me.lblMF01.Text = str1
                    Me.lblMF01.Visible = True
                Else

                    Me.lblMF01.Visible = False

                End If

        End Select

    End Sub

    Private Sub chkInclIntStdNMF_CheckedChanged(sender As Object, e As EventArgs) Handles chkInclIntStdNMF.CheckedChanged

        If boolHold Then
        Else
            If CheckMF(Me.chkInclIntStdNMF) Then
                Call UpdateChkData01()
                Call UpdateChkData(False, Me.chkInclIntStdNMF, "BOOLINCLINTSTDNMF")
            End If
        End If

    End Sub

    Sub UpdateChkData01()

        Dim idC As Int32 = idCGet()

        Select Case idC
            Case 17
                'Me.lblMF01.Visible = False
                Call DolblMF()
            Case Else
                If Me.chkInclIntStdNMF.Checked Then
                    Me.chkCalcIntStdNMF.Enabled = True
                    Me.lblMF01.Visible = True
                Else
                    Me.chkCalcIntStdNMF.Checked = False
                    Me.chkCalcIntStdNMF.Enabled = False
                    Me.lblMF01.Visible = False
                End If
        End Select



    End Sub

    Private Sub chkCalcIntStdNMF_CheckedChanged(sender As Object, e As EventArgs) Handles chkCalcIntStdNMF.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkCalcIntStdNMF, "BOOLCALCINTSTDNMF")
            'Call DoMFTable()
            'Call UpdateCalc(True)
        End If

    End Sub

    Sub UpdateCalc(boolDoLabel As Boolean)

        Dim str1 As String
        'If Me.chkCalcIntStdNMF.Checked Then
        '    str1 = "Include Matrix Factor columns for Analyte and Int Std"
        'Else
        '    str1 = "Include Peak Area Ratio columns for Recovery and Post Extraction Spiking"
        'End If
        'Me.chkInclMFCols.Text = str1

        If boolHold Then
        Else

            If boolDoLabel Then

                Dim boolPos As Boolean
                boolPos = Me.rbPosLeg.Checked

                Dim strLegTitle As String = DoLegendThingsMatrix(boolPos, "Title")
                Dim strLegNum As String = DoLegendThingsMatrix(boolPos, "Num")
                Dim strLegDen As String = DoLegendThingsMatrix(boolPos, "Den")

                Me.CHARTITLELEG.Text = strLegTitle
                Me.CHARNUMLEG.Text = strLegNum
                Me.CHARDENLEG.Text = strLegDen

                Call UpdateTBData(Me.CHARTITLELEG, "CHARTITLELEG")
                Call UpdateTBData(Me.CHARNUMLEG, "CHARNUMLEG")
                Call UpdateTBData(Me.CHARDENLEG, "CHARDENLEG")

            End If
        End If

    End Sub

    Private Sub CHARCARRYOVERLABEL_Validated(sender As Object, e As EventArgs) Handles CHARCARRYOVERLABEL.Validated

        If boolHold Then
        Else
            Call UpdateTBData(Me.CHARCARRYOVERLABEL, "CHARCARRYOVERLABEL")
        End If

    End Sub

    Private Sub CHARCARRYOVERLABEL_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles CHARCARRYOVERLABEL.Validating

        Dim strM As String
        Dim str1 As String

        str1 = Me.CHARCARRYOVERLABEL.Text

        'cannot be blank
        If Len(str1) = 0 Then
            strM = "Carryover Label entry cannot be zero-length."
            strM = strM & ChrW(10) & ChrW(10) & "Default of 'Blank' will be entered."
            MsgBox(strM, vbInformation, "Invalid entry...")
            Me.CHARCARRYOVERLABEL.Text = "Blank"
            e.Cancel = True
            GoTo end1
        End If

        Dim dgv As DataGridView = Me.dgvReportTables
        Dim int1 As Short = dgv.CurrentCell.RowIndex
        Dim str2 As String = dgv("CHARHEADINGTEXT", int1).Value

        Dim strMod As String = "Advanced Report Table Configuration - " & str2
        Dim strSource As String = "Carryover Label"

        If CheckColLenEx(str1, 255, strMod, strSource) Then
            e.Cancel = True
        End If

end1:

    End Sub

   
    Private Sub CHARNUMLEG_TextChanged(sender As Object, e As EventArgs) Handles CHARNUMLEG.TextChanged

    End Sub

    Function CheckMF(chk As CheckBox)

        CheckMF = True

        Dim idC As Int32 = idCGet()

        Select Case idC

            Case 17

            Case Else


                If Me.chkMFTable.Checked And boolHold = False And Me.gbMatrixFactor.Visible = True Then

                    Dim strM As String
                    Dim bool1 As Boolean = Me.chkInclMFCols.Checked
                    Dim bool2 As Boolean = Me.chkInclIntStdNMF.Checked

                    If bool1 = False And bool2 = False Then
                        strM = "If 'Show table as Matrix Factor' is checked, then one of these two checkboxes must be checked."
                        MsgBox(strM, vbInformation, "Invalid action...")
                        chk.Checked = True
                    End If

                End If

        End Select


    End Function

    Private Sub CHARCARRYOVERLABEL_TextChanged(sender As Object, e As EventArgs) Handles CHARCARRYOVERLABEL.TextChanged

    End Sub

    Private Sub NUMPRECCRITLOTS_TextChanged(sender As Object, e As EventArgs) Handles NUMPRECCRITLOTS.TextChanged

    End Sub

    Private Sub NUMPRECCRITLOTS_Validated(sender As Object, e As EventArgs) Handles NUMPRECCRITLOTS.Validated

        If boolHold Then
        Else
            Call UpdateNumeric(Me.NUMPRECCRITLOTS, "NUMPRECCRITLOTS")
        End If

    End Sub

    Private Sub NUMPRECCRITLOTS_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles NUMPRECCRITLOTS.Validating

        'entry must be numeric >= 0

        Dim var1
        Dim strM As String
        Dim boolE As Boolean = True

        strM = "Entry must be numeric >= 0"

        var1 = Me.NUMPRECCRITLOTS.Text

        If Len(var1) = 0 Then
            GoTo end1
        End If

        If IsNumeric(var1) Then
        Else
            GoTo end1
        End If

        If var1 < 0 Then
            GoTo end1
        End If

        boolE = False
end1:

        If boolE Then
            MsgBox(strM, vbInformation, "Invalid entry...")
            e.Cancel = True
        End If

    End Sub

    Private Sub chkBOOLREGRULOQ_CheckedChanged(sender As Object, e As EventArgs) Handles chkBOOLREGRULOQ.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkBOOLREGRULOQ, "BOOLREGRULOQ")
        End If

    End Sub

    Private Sub rbINTQCLEVELGROUPLevel_CheckedChanged(sender As Object, e As EventArgs) Handles rbINTQCLEVELGROUPLevel.CheckedChanged
        If boolHold Then
        Else
            Call UpdateNumeric(Me.rbINTQCLEVELGROUPLevel, "INTQCLEVELGROUP")
        End If
    End Sub

    Private Sub rbINTQCLEVELGROUPNomConc_CheckedChanged(sender As Object, e As EventArgs) Handles rbINTQCLEVELGROUPNomConc.CheckedChanged
        If boolHold Then
        Else
            Call UpdateNumeric(Me.rbINTQCLEVELGROUPNomConc, "INTQCLEVELGROUP")
        End If
    End Sub

    Private Sub rbINTQCLEVELGROUPQCLabel_CheckedChanged(sender As Object, e As EventArgs) Handles rbINTQCLEVELGROUPQCLabel.CheckedChanged
        If boolHold Then
        Else
            Call UpdateNumeric(Me.rbINTQCLEVELGROUPQCLabel, "INTQCLEVELGROUP")
        End If
    End Sub

    Private Sub cbxSampleSAD5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSampleSAD5.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleSAD5)

    End Sub

    Private Sub cbxSampleSAD6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSampleSAD6.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleSAD6)

    End Sub

    Private Sub cbxSampleS5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSampleS5.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleS5)

    End Sub

    Private Sub cbxSampleS6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSampleS6.SelectedIndexChanged

        Call UpdateGroupSortData(Me.cbxSampleS6)

    End Sub

    Private Sub dgvAnalytes_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAnalytes.CellContentClick

    End Sub

    Private Sub chkBOOLCONCCOMMENTS_CheckedChanged(sender As Object, e As EventArgs) Handles chkBOOLCONCCOMMENTS.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkBOOLCONCCOMMENTS, "BOOLCONCCOMMENTS")
        End If

    End Sub

    Private Sub rbDontShowBQL_CheckedChanged(sender As Object, e As EventArgs) Handles rbDontShowBQL.CheckedChanged

    End Sub

    Private Sub chkBOOLADHOCSTABCOMPCOLUMNS_CheckedChanged(sender As Object, e As EventArgs) Handles chkBOOLADHOCSTABCOMPCOLUMNS.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(False, Me.chkBOOLADHOCSTABCOMPCOLUMNS, "BOOLADHOCSTABCOMPCOLUMNS")
        End If

    End Sub

    Private Sub rbOld_CheckedChanged(sender As Object, e As EventArgs) Handles rbOld.CheckedChanged

        'Call DoLegendThings()
        If boolHold Then
        Else
            Call DoLegendThings()
            Call UpdateChkData(False, Me.chkCalcIntStdNMF, "BOOLCALCINTSTDNMF")
        End If

    End Sub

    Private Sub rbNew_CheckedChanged(sender As Object, e As EventArgs) Handles rbNew.CheckedChanged

        'Call DoLegendThings()
        If boolHold Then
        Else
            Call DoLegendThings()
            Call UpdateChkData(False, Me.chkCalcIntStdNMF, "BOOLCALCINTSTDNMF")
        End If

    End Sub

    Private Sub chkInjCol_CheckedChanged(sender As Object, e As EventArgs) Handles chkInjCol.CheckedChanged

        If boolHold Then
        Else
            If Me.chkInjCol.Checked Then
                Me.chkBOOLCONCCOMMENTS.Checked = True
            Else
                Me.chkBOOLCONCCOMMENTS.Checked = False
            End If

        End If

    End Sub

    Sub DoInjCol()

        Try

            Dim idC As Int32 = idCGet()

            Select Case idC
                Case 35

                    If BOOLCONCCOMMENTS Then
                        Me.chkInjCol.Checked = True
                    Else
                        Me.chkInjCol.Checked = False
                    End If
            End Select


        Catch ex As Exception

        End Try

    End Sub

   
    Private Sub rbNA_CheckedChanged(sender As Object, e As EventArgs) Handles rbNA.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbNA, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbProcess_CheckedChanged(sender As Object, e As EventArgs) Handles rbProcess.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbProcess, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbBenchTop_CheckedChanged(sender As Object, e As EventArgs) Handles rbBenchTop.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbBenchTop, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbFT_CheckedChanged(sender As Object, e As EventArgs) Handles rbFT.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbFT, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbLT_CheckedChanged(sender As Object, e As EventArgs) Handles rbLT.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbLT, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbReinjection_CheckedChanged(sender As Object, e As EventArgs) Handles rbReinjection.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbReinjection, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbBlood_CheckedChanged(sender As Object, e As EventArgs) Handles rbBlood.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbBlood, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbStockSolution_CheckedChanged(sender As Object, e As EventArgs) Handles rbStockSolution.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbStockSolution, "BOOLSTATSNR")
        End If

    End Sub

   

    Private Sub txtStabilityNotes_Validated(sender As Object, e As EventArgs) Handles txtStabilityNotes.Validated

        If boolHold Then
        Else
            Call UpdateTBData(Me.txtStabilityNotes, "CHARCARRYOVERLABEL")
        End If

    End Sub

    Private Sub INTNUMBEROFCYCLES_TextChanged(sender As Object, e As EventArgs) Handles INTNUMBEROFCYCLES.TextChanged

        '20181112 LEE:
        If boolHold Then
        Else
            Me.lblRemember.Visible = True
        End If

    End Sub

    Private Sub CHARTIMEPERIOD_TextChanged(sender As Object, e As EventArgs) Handles CHARTIMEPERIOD.TextChanged

        '20181112 LEE:
        If boolHold Then
        Else
            Me.lblRemember.Visible = True
        End If

    End Sub

    Private Sub CHARTIMEFRAME_TextChanged(sender As Object, e As EventArgs) Handles CHARTIMEFRAME.TextChanged

        '20181112 LEE:
        If boolHold Then
        Else
            Me.lblRemember.Visible = True
        End If

    End Sub

    Private Sub CHARPERIODTEMP_TextChanged(sender As Object, e As EventArgs) Handles CHARPERIODTEMP.TextChanged

        '20181112 LEE:
        If boolHold Then
        Else
            Me.lblRemember.Visible = True
        End If

    End Sub


    Private Sub rbSpiking_CheckedChanged(sender As Object, e As EventArgs) Handles rbSpiking.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbSpiking, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub chkRTC_CalStd_Acc_CheckedChanged(sender As Object, e As EventArgs) Handles chkRTC_CalStd_Acc.CheckedChanged

        '20181220 LEE:
        'BOOLCSREPORTACCVALUES will be used for Do Calculations True/False
        If boolHold Then
        Else
            'Call UpdateChkData(True, Me.rbRTC_CalStd_Acc, "BOOLCSREPORTACCVALUES")
            Call UpdateChkData(False, Me.chkRTC_CalStd_Acc, "BOOLCSREPORTACCVALUES")
        End If

    End Sub

    Private Sub rbAutosampler_CheckedChanged(sender As Object, e As EventArgs) Handles rbAutosampler.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbAutosampler, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub rbBatchReinjection_CheckedChanged(sender As Object, e As EventArgs) Handles rbBatchReinjection.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbBatchReinjection, "BOOLSTATSNR")
        End If

    End Sub

    Private Sub CHARSTABILITYPERIOD_TextChanged(sender As Object, e As EventArgs) Handles CHARSTABILITYPERIOD.TextChanged

    End Sub

    Private Sub txtStabilityNotes_TextChanged(sender As Object, e As EventArgs) Handles txtStabilityNotes.TextChanged

    End Sub

    Private Sub rbDilution_CheckedChanged(sender As Object, e As EventArgs) Handles rbDilution.CheckedChanged

        If boolHold Then
        Else
            Call UpdateChkData(True, Me.rbDilution, "BOOLSTATSNR")
        End If

    End Sub

End Class