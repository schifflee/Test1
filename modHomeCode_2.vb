Option Compare Text

Imports System
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Word

Module modHomeCode_2

    Sub ToolTipSet()

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

            'General Buttons
            toolTip1.SetToolTip(frmH.cmdEdit, "Change to editing mode")
            toolTip1.SetToolTip(frmH.cmdSave, "Save all changes")
            toolTip1.SetToolTip(frmH.cmdCancel, "Cancel unsaved changes")
            toolTip1.SetToolTip(frmH.cmdExit, "Exit Report Writer")
            toolTip1.SetToolTip(frmH.cmdSymbol, "Copy non-keyboard characters for text entry.")
            toolTip1.SetToolTip(frmH.cmdHomeCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.llblAssignedSamples, "Click here to go to the Assign Samples page to populate empty tables.")        'TAB 1: Choose Study & Template
            toolTip1.SetToolTip(frmH.lblHome, "Choose Watson Study and StudyDoc Template  to apply")
            toolTip1.SetToolTip(frmH.cmdUpdateProject, "Select/switch Watson study")
            toolTip1.SetToolTip(frmH.cmdCreateReportTitle2, "Construct a Report Title from Top Level Data information")
            toolTip1.SetToolTip(frmH.cmdShowOutstanding, "Show any issues with last generated report")
            toolTip1.SetToolTip(frmH.cmdUpdateSummaryInfo, "Update information that has been entered in other portions of StudyDoc")

            'Grid: Configured report
            frmH.dgvReports.Columns.Item("CHARREPORTNUMBER").ToolTipText = "Enter/Edit Report Number"
            frmH.dgvReports.Columns.Item("CHARREPORTTITLE").ToolTipText = "Enter/Edit Report Title (or use Create Report Title button above)"
            frmH.dgvReports.Columns.Item("DTREPORTDRAFTISSUEDATE").ToolTipText = "Enter/Edit Date of this draft"
            frmH.dgvReports.Columns.Item("DTREPORTFINALISSUEDATE").ToolTipText = "Enter/Edit Date Report is being issued"

            'TAB 2: Add/Edit Top Level Data
            toolTip1.SetToolTip(frmH.cmdDataCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.cbxSubmittedTo, "Select Corporate Address from pre-set list (set in Report Writer Administration)")
            toolTip1.SetToolTip(frmH.cbxInSupportOf, "Select Corporate Address from pre-set list (set in Report Writer Administration)")
            toolTip1.SetToolTip(frmH.cbxSubmittedBy, "Select Corporate Address from pre-set list (set in Report Writer Administration)")

            'Subtabs
            frmH.tabData.ShowToolTips = True
            frmH.tabData1.ToolTipText = "Enter Top-Level Study Information"
            frmH.tabData2.ToolTipText = "Set global options for this Study"
            frmH.tabData3.ToolTipText = "Add and Set Template specific variables"
            'Study Information Subtab
            'row-based Tooltips are outlined in dgvStudyConfig_CellMouseEnter routine

            'Round Conventions Subtab
            toolTip1.SetToolTip(frmH.gbRound5, "Applies to all Raw and Calculated data configured for rounding" & _
                                vbCrLf & "(either significant figures or decimals)")
            toolTip1.SetToolTip(frmH.rbRoundFiveEven, "Round 5 to the nearest even number" & _
                                vbCrLf & "(e.g.  12.25 rounds to 12.2; 17.55 rounds to 17.6)." & ChrW(10) & "Watson LIMS" & ChrW(8482) & " uses this convention")
            toolTip1.SetToolTip(frmH.rbRoundFiveAway, "Round 5 to the next higher number (when above zero) or the next lower number" & _
                                vbCrLf & "(when below zero).   (e.g.  12.25 rounds to 12.3,   17.55 rounds to 17.6, " & _
                                vbCrLf & "-12.25 rounds to -12.3)." & ChrW(10) & "Excel ROUND function uses this convention.")


            toolTip1.SetToolTip(frmH.gbCritPrecision, "Governs how plus/minus percentage criteria (e.g. +/-15%) are applied.")
            toolTip1.SetToolTip(frmH.rbCritFullPrec, "Keep full precision on the criteria for comparison" & _
                                vbCrLf & "(e.g. +287.5 for 250 + 15%), meaning 288 is out of criteria.")
            toolTip1.SetToolTip(frmH.rbCritRounded, "Rounds the  full  the criteria for comparison " & _
                                vbCrLf & "(e.g. +288 for 250 + 15%, meaning 288 is within criteria.")
            toolTip1.SetToolTip(frmH.gbMeanComp, "Governs whether to keep precision when calculating averages" & _
                                vbCrLf & "or differences between means (e.g. Recovery Calculations).")
            toolTip1.SetToolTip(frmH.rbMeanFullPrec, "Use full precision during calculation")
            toolTip1.SetToolTip(frmH.rbMeanRounded, "Round each mean before making comparison")

            'TAB 3: Review Analytical Runs
            toolTip1.SetToolTip(frmH.cmdAnaRunSumCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.cmdViewAnalyticalRuns1, "Inspect list of all analytical runs in study")

            'BOOLALLAR chkAll
            'BOOLACCAR chkAccepted
            'BOOLREJAR chkRejected
            'BOOLREGRAR chkRegrPerformed
            'BOOLNOREGRAR chkNoRegrPerformed
            'BOOLINCLPSAE chkPSAE
            toolTip1.SetToolTip(frmH.chkAll, "Include all analytical runs")
            toolTip1.SetToolTip(frmH.chkAccepted, "Include Accepted analytical runs")
            toolTip1.SetToolTip(frmH.chkRejected, "Include Rejected analytical runs")
            toolTip1.SetToolTip(frmH.chkRegrPerformed, "Include 'Regression Performed' analytical runs")
            toolTip1.SetToolTip(frmH.chkNoRegrPerformed, "Include 'NO Regression Performed' analytical runs")
            toolTip1.SetToolTip(frmH.chkPSAE, "Include Pre Study Assay Evaluation (PSAE) analytical runs")

            'Review Watson Runs table
            frmH.dgvAnalyticalRunSummary.Columns(0).ToolTipText = "A: Include in Run Summary Table"
            frmH.dgvAnalyticalRunSummary.Columns(1).ToolTipText = "B: Include in Regression Table"

            'TAB 4: Summary Table - Method Validation and Study Information
            toolTip1.SetToolTip(frmH.cmdOrderSummaryTable, "Re-order the rows (after changing order #'s")
            toolTip1.SetToolTip(frmH.lblSummaryTable, "Decide what information from the StudyDoc Method Validation study should be reported in the Summary Table")
            toolTip1.SetToolTip(frmH.cmdResetSummaryTable, "Undo unsaved changes (this tab only)")
            'Table columns - this isn't working
            frmH.dgvSummaryData.Columns.Item("boolI").ToolTipText = "A: Include in the Summary Table for the Validated Method"
            frmH.dgvSummaryData.Columns.Item("charValue").ToolTipText = "Value: Read-only (from Watson, or StudyDoc - Review Validated Method section"

            'TAB 5: Choose/Edit Word Template
            toolTip1.SetToolTip(frmH.cmdCancelReportStatements, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.cmdRefreshStatements, "Show all word templates available")
            toolTip1.SetToolTip(frmH.cmdOpenReportStatements, "Create/Edit any word template")
            toolTip1.SetToolTip(frmH.lblWordStatements, "Choose Word template for report.  Please ensure it matches the study")

            'TAB 6: Configure Report Tables
            toolTip1.SetToolTip(frmH.cmdRTConfigCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.lblReportTableConfiguration, "Choose and configure tables to be included in table section of report")
            toolTip1.SetToolTip(frmH.cmdOrderReportTableConfig, "Re-order the rows (after changing the order #'s)")
            toolTip1.SetToolTip(frmH.cmdCreateTable, "Generate the report currently selected")
            toolTip1.SetToolTip(frmH.cmdAssignSamples, "Manually choose/review which samples get included in each table")
            toolTip1.SetToolTip(frmH.cmdAdvancedTable, "Select options (statistical, etc.) for each table)")
            toolTip1.SetToolTip(frmH.cmdDuplicateTables, "Create new table with same properties as currently selected table")
            toolTip1.SetToolTip(frmH.cmdOutliers, "Open Outlier Evaluation screen for appropriate tables")
            toolTip1.SetToolTip(frmH.cmdViewAnalRuns, "Inspect list of all analytical runs in study")
            toolTip1.SetToolTip(frmH.cmdSelect, "Select Analyte column(s) for a group of cells")
            toolTip1.SetToolTip(frmH.cmdImportTables, "Create new table based on a table from another StudyDoc template")
            toolTip1.SetToolTip(frmH.cmdRTCDown, "Move currently selected row down 1 row")
            toolTip1.SetToolTip(frmH.cmdRTCUp, "Move currently selected row up 1 row")
            'Do Grid Columns    
            frmH.dgvReportTableConfiguration.Columns.Item("boolRequiresSampleAssignment").ToolTipText = "A: Select if table requires samples to be manually assigned"
            frmH.dgvReportTableConfiguration.Columns.Item("charPageOrientation").ToolTipText = "P/L: Orient table (P=Portrait, L=Landscape)"
            frmH.dgvReportTableConfiguration.Columns.Item("boolInclude").ToolTipText = "Include table in Tables Section"
            frmH.dgvReportTableConfiguration.Columns.Item("charFCID").ToolTipText = "FC ID: Field Code ID (for adding individual table to report)"
            frmH.dgvReportTableConfiguration.Columns.Item("boolPlaceHolder").ToolTipText = "If selected, only placeholder (vs full table) will be included in report"
            frmH.dgvReportTableConfiguration.Columns.Item("charFCID").ToolTipText = "FC ID: Field Code ID (for adding individual table to report)"
            frmH.dgvReportTableConfiguration.Columns.Item("charFCID").ToolTipText = "FC ID: Field Code ID (for adding individual table to report)"

            'TAB 7: Configure Column Headings
            toolTip1.SetToolTip(frmH.cmdRTHeaderConfigCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.Label35, "Change column headings on the report")
            toolTip1.SetToolTip(frmH.lbldgvReportTables, "Select table (to change its column headers)")
            'Do Grid Columns (Report Tables)
            '   NDL: looks like this grid is used for a lot more than just this screen (many columns).  I will not tooltip it for now.
            'Do Grid Columns (Table Column Headers)
            frmH.dgvReportTableHeaderConfig.Columns.Item("charColumnLabel").ToolTipText = "Header suggested by StudyDoc or Watson"
            frmH.dgvReportTableHeaderConfig.Columns.Item("charUserLabel").ToolTipText = "Enter preferred header for report"
            frmH.dgvReportTableHeaderConfig.Columns.Item("intOrder").ToolTipText = "Set column order on report"
            frmH.dgvReportTableHeaderConfig.Columns.Item("boolInclude").ToolTipText = "Include this column on report"

            'TAB 8: Analytical Reference Std
            toolTip1.SetToolTip(frmH.cmdAnalRefCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.Label9, "Inspect/Add Analytical Reference Standard Information")
            toolTip1.SetToolTip(frmH.cmdAddAnalyte, "Add standard (that is not in Watson) to the study" _
                                & vbCrLf & "(e.g. co-administered compounds)")
            toolTip1.SetToolTip(frmH.cmdAddRepAnalyte, "Add replicate of existing study analyte" _
                                & vbCrLf & "(e.g. for an additional lot of analyte)")
            toolTip1.SetToolTip(frmH.cmdCopyRepAnalyte, "Copy StudyDoc information from one analyte/standard to another")
            toolTip1.SetToolTip(frmH.cmdDeleteRepAnalyte, "Delete a standard or analyte that was added via StudyDoc")
            frmH.dgvCompanyAnalRef.Columns.Item("boolInclude").ToolTipText = "A: Include this item in the report"

            'TAB 9: Add/Edit Contributors
            toolTip1.SetToolTip(frmH.lblGlobalConfiguration, "Select contributors to the current report")
            toolTip1.SetToolTip(frmH.cmdCPCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.cmdCPAdd, "Add a contributor")
            toolTip1.SetToolTip(frmH.cmdCPDelete, "Delete a contributor")
            toolTip1.SetToolTip(frmH.cmdReplacePersonnel, "Replace this table of contributors with a table from another study")
            frmH.dgvContributingPersonnel.Columns.Item("boolIncludeSOTP").ToolTipText = "A: Include on Contributing Personnel page"
            frmH.dgvContributingPersonnel.Columns.Item("charCPName").ToolTipText = "Choose name (from Global User Accounts + imported lists)" '
            frmH.dgvContributingPersonnel.Columns.Item("charCPSuffix").ToolTipText = "Choose name suffix (from RW Administration)"
            frmH.dgvContributingPersonnel.Columns.Item("charCPDegree").ToolTipText = "Choose educational degree (from RW Administration)"
            frmH.dgvContributingPersonnel.Columns.Item("charCPPrefix").ToolTipText = "Choose Name Prefix (from RW Administration)"
            frmH.dgvContributingPersonnel.Columns.Item("charCPRole").ToolTipText = "Choose Personnel Role (from RW Administration)"
            frmH.dgvContributingPersonnel.Columns.Item("charCPTitle").ToolTipText = "Choose Personnel Title (from RW Administration)"
            frmH.dgvContributingPersonnel.Columns.Item("intOrder").ToolTipText = "B: Edit numbers to change order of contributors on Contributing Personnel Page"

            'TAB 10: Review Validated Method
            toolTip1.SetToolTip(frmH.Label17, "Assign the method validation study and review its values")
            toolTip1.SetToolTip(frmH.cmdMethValUpdate, "(Re)Apply all Validated Method values from the method validation study selected, to this study")
            toolTip1.SetToolTip(frmH.cmdMethValReset, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.cmdMethValExecute, "Apply all Validated Method values from the method validation study selected to this study (above)")
            'Doesn't work:   toolTip1.SetToolTip(frmH.cbxArchivedMDB, "Choose the StudyDoc-configured Validation Study")

            'TAB 11: QA Event Table
            toolTip1.SetToolTip(frmH.Label32, "Create/maintain the table of QA Events")
            toolTip1.SetToolTip(frmH.cmdInsertQAEvent, "Insert row below selected row")
            toolTip1.SetToolTip(frmH.cmdQACancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.cmdDeleteQAEvent, "Delete selected row")

            'TAB 12: Sample Receipt Records
            toolTip1.SetToolTip(frmH.cmdSRecCancel, "Undo unsaved changes (this tab only)")
            toolTip1.SetToolTip(frmH.cmdInsertSRec, "Insert row below selected row")
            toolTip1.SetToolTip(frmH.cmdDeletSRec, "Delete selected row")
            toolTip1.SetToolTip(frmH.txtSRecTotalReportWatson, "Total Sample Count from Watson reports")
            'Do Grid Columns (Sample Receipt Records)
            frmH.dgvSampleReceipt.Columns.Item("BoolU").ToolTipText = "A: Include this sample count in ""total number of samples received"" sum calculation for Report"
            frmH.dgvSampleReceipt.Columns.Item("DTSHIPMENTRECEIVED").ToolTipText = "Enter Date Received"
            frmH.dgvSampleReceipt.Columns.Item("NUMSAMPLENUMBER").ToolTipText = "Enter Sample Count"
            frmH.dgvSampleReceipt.Columns.Item("CHARSTORAGETEMP").ToolTipText = "Enter Storage Temperature"
            frmH.dgvSampleReceipt.Columns.Item("CHARCONDITION").ToolTipText = "Enter Sample Condition (e.g. ""Adequate"")"
            frmH.dgvSampleReceipt.Columns.Item("CHARSOURCE").ToolTipText = "Enter Source of sample"

            str1 = "Studies whose Issue Date is blank."
            toolTip1.SetToolTip(frmH.optStudyDocOpen, str1)

            str1 = "Studies whose Issue Date is not blank."
            toolTip1.SetToolTip(frmH.optStudyDocClosed, str1)

            str1 = "Denotes Lock status of last generated Final Report."
            str1 = str1 & ChrW(10) & "If checked, user may still create tables and sections, but not a Final Report."
            toolTip1.SetToolTip(frmH.chkLockFinalReport, str1)

            'optStudyDocOpen()
            'optStudyDocClosed()

        Catch ex As Exception

        End Try


    End Sub

    Sub ConfigLockFinalReport()

        'first determine if panLockFinalReport is to be shown
        If gboolER Then
            frmH.panLockFinalReport.Visible = True
        Else
            frmH.panLockFinalReport.Visible = False
        End If

        'determine check status

        Dim strF As String
        Dim strS As String
        Dim intL As Short

        strF = "CHARREPORTTYPE = 'Final Report' AND ID_TBLSTUDIES = " & id_tblStudies
        strS = "UPSIZE_TS DESC"

        Dim rows() As DataRow = tblFinalReport.Select(strF, strS, DataViewRowState.CurrentRows)
        Dim boolF As Boolean
        boolF = boolFormLoad
        boolFormLoad = True
        If rows.Length = 0 Then
            frmH.chkLockFinalReport.Checked = False
            frmH.lblFinalReportLockedDate.Text = "NA"
        Else
            intL = NZ(rows(0).Item("BOOLLOCKED"), 0)
            If intL = 0 Then
                frmH.chkLockFinalReport.Checked = False
            Else
                frmH.chkLockFinalReport.Checked = True
            End If

            'enter label information
            Try
                Dim dt As Date = rows(0).Item("UPSIZE_TS")
                Dim str1 As String = LTextDateFormat & " HH:mm:ss tt"
                Dim strDt As String = Format(dt, str1)
                frmH.lblFinalReportLockedDate.Text = strDt

            Catch ex As Exception
                frmH.lblFinalReportLockedDate.Text = "NA"
            End Try

        End If
        boolFormLoad = boolF

        If frmH.chkLockFinalReport.Checked Then
            BOOLFINALREPORTLOCKED = True
        Else
            BOOLFINALREPORTLOCKED = False
        End If

    End Sub

    Sub SaveFinalReport()

        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblFinalReport)

            If boolGuWuOracle Then
                Try
                    'ta_tblFinalReport.Update(tblFinalReport)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLREPORTSTATEMENTS.Merge('ds2005.TBLREPORTSTATEMENTS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_TBLFINALREPORTAcc.Update(tblFinalReport)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTSTATEMENTS.Merge('ds2005Acc.TBLREPORTSTATEMENTS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_TBLFINALREPORTSQLServer.Update(tblFinalReport)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTSTATEMENTS.Merge('ds2005Acc.TBLREPORTSTATEMENTS, True)
                End Try
            End If

        End If

    End Sub


    Sub PositionAbort(ByVal boolOpen As Boolean)

        Exit Sub

        Try

            If boolOpen Then
                frmAbort.Show()
                frmAbort.Left = frmH.lblProgress.Left + frmH.lblProgress.Width + 10
                frmAbort.Top = (frmH.lblProgress.Top + frmH.lblProgress.Height) - frmAbort.Height
            Else
                frmAbort.Visible = False
                frmAbort.Dispose()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Sub cbxRBSFilterPopulate()
        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim rows() As DataRow
        Dim strS As String
        Dim strF As String

        tbl = tblConfigBodySections
        dv = tbl.DefaultView
        Dim tbl1 As System.Data.DataTable = dv.ToTable(True, "NUMCOMPANY")
        strS = "NUMCOMPANY ASC"
        strF = "NUMCOMPANY > 0"
        rows = tbl1.Select(strF, strS)
        int1 = rows.Length

        frmH.cbxRBSFilter.Items.Add("[None]")
        frmH.cbxRBSFilter.Items.Add("Company")

        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("NUMCOMPANY"), 0)
            frmH.cbxRBSFilter.Items.Add(str1)
        Next

        'choose first item
        frmH.cbxRBSFilter.SelectedIndex = 0

        'populate public tblreportcompanies
        Dim col2 As New DataColumn
        col2.ColumnName = "ID"
        tblReportCompanies.Columns.Add(col2)
        For Count1 = 0 To int1 - 1 + 2
            Dim row1 As DataRow = tblReportCompanies.NewRow()
            row1.BeginEdit()
            str1 = frmH.cbxRBSFilter.Items(Count1)
            row1.Item("ID") = str1
            row1.EndEdit()
            tblReportCompanies.Rows.Add(row1)
        Next


    End Sub

    Sub cbxRBSTypeFilterPopulate()

        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim rows() As DataRow
        Dim strF As String

        tbl = tblConfigReportType
        strF = "ID_TBLCONFIGREPORTTYPE < 1000 AND BOOLINCLUDE = -1"
        rows = tbl.Select(strF)
        int1 = rows.Length

        frmH.cbxRBSTypeFilter.Items.Add("[None]")

        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("CHARREPORTTYPE"), "Sample Analysis")
            frmH.cbxRBSTypeFilter.Items.Add(str1)
        Next

        'choose first item
        frmH.cbxRBSTypeFilter.SelectedIndex = 0

    End Sub

    Sub cbxFilterPopulate()

        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim dv As System.Data.DataView

        tbl = tblStudies
        dv = tbl.DefaultView
        Dim tbl1 As System.Data.DataTable = dv.ToTable(True, "charCust")
        Dim int1 As Short
        Dim Count1 As Short
        int1 = tbl1.Rows.Count
        frmH.cbxFilter.Items.Clear()
        frmH.cbxFilter.Items.Add("All")
        For Count1 = 0 To int1 - 1
            frmH.cbxFilter.Items.Add(tbl1.Rows.Item(Count1).Item("charCust"))
        Next

    End Sub

    Sub ReportsHomeInitialize()

        Dim dgt As DataGridTableStyle
        Dim ts1 As New DataGridTableStyle
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim boolRO As Boolean
        Dim dgv As DataGridView
        Dim intC As Short
        Dim strN As String
        Dim tbl As System.Data.DataTable
        Dim int1 As Short
        Dim var1
        Dim strF As String
        Dim strS As String

        dgv = frmH.dgvReports
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidth = 25
        dgv.AllowUserToOrderColumns = False
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        'dgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        dgv.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        tbl = tblReports
        intC = tblReports.Columns.Count
        'filter dgvReports
        strF = "id_tblStudies = 0"
        strS = "id_tblReports ASC"
        Dim dv As System.Data.DataView = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)
        dv.RowFilter = "id_tblStudies = 0"
        dv.AllowDelete = False
        dv.AllowNew = False
        dv.AllowEdit = True
        dgv.DataSource = dv

        For Count1 = 0 To intC - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            'dgv.Columns.item(Count1).MinimumWidth = 75
            dgv.Columns.Item(Count1).DisplayIndex = intC - 1
        Next

        For Count1 = 0 To intC - 1
            strN = tbl.Columns.Item(Count1).ColumnName
            str1 = "yaya"
            Select Case strN
                Case "CHARREPORTNUMBER"
                    str1 = "CHARREPORTNUMBER"
                    str2 = "Report Number"
                    boolRO = False
                    int1 = 0
                    var1 = DataGridViewContentAlignment.BottomLeft
                    dgv.Columns.Item(str1).MinimumWidth = 75
                Case "CHARREPORTTITLE"
                    str1 = "CHARREPORTTITLE"
                    str2 = "Report Title"
                    boolRO = False
                    dgv.Columns.Item(str1).MinimumWidth = 300
                    int1 = 1
                    var1 = DataGridViewContentAlignment.BottomLeft
                Case "DTREPORTDRAFTISSUEDATE"
                    str1 = "DTREPORTDRAFTISSUEDATE"
                    str2 = "Draft Date"
                    boolRO = False
                    int1 = 2
                    var1 = DataGridViewContentAlignment.BottomCenter
                Case "DTREPORTFINALISSUEDATE"
                    str1 = "DTREPORTFINALISSUEDATE"
                    str2 = "Issue Date"
                    boolRO = False
                    int1 = 3
                    var1 = DataGridViewContentAlignment.BottomCenter
                Case "CHARREPORTTYPE"
                    str1 = "CHARREPORTTYPE"
                    str2 = "Study Type *"
                    boolRO = False
                    int1 = 4
                    var1 = DataGridViewContentAlignment.BottomLeft
                Case "CHARREPORTTEMPLATE"
                    str1 = "CHARREPORTTEMPLATE"
                    str2 = "MSWord Doc Template *"
                    boolRO = False
                    int1 = 5
                    var1 = DataGridViewContentAlignment.BottomLeft
                    'Case "ID_TBLCONFIGREPORTTYPE"
                    '    str1 = "ID_TBLCONFIGREPORTTYPE"
                    '    str2 = "ID_TBLCONFIGREPORTTYPE"
                    '    boolRO = True
                    '    int1 = 6
                    '    var1 = DataGridViewContentAlignment.BottomCenter
            End Select

            If StrComp(strN, str1, CompareMethod.Text) = 0 Then
                dgv.Columns.Item(strN).Visible = True
                dgv.Columns.Item(strN).HeaderText = str2
                dgv.Columns.Item(strN).ReadOnly = boolRO
                dgv.Columns.Item(strN).DisplayIndex = int1
                dgv.Columns.Item(strN).DefaultCellStyle.Alignment = var1
            End If

            If StrComp(strN, "CHARREPORTTEMPLATE", CompareMethod.Text) = 0 Then 'hide
                dgv.Columns.Item(strN).Visible = False
            End If
        Next

        str2 = LDateFormat
        dgv.Columns.Item("DTREPORTDRAFTISSUEDATE").DefaultCellStyle.Format = str2
        dgv.Columns.Item("DTREPORTFINALISSUEDATE").DefaultCellStyle.Format = str2

        dgv.AutoResizeRows()



    End Sub

    Sub TempReportsOrder()

        Dim dgt As DataGridTableStyle
        Dim ts1 As New DataGridTableStyle
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim boolRO As Boolean
        Dim dgv As DataGridView
        Dim intC As Short
        Dim strN As String
        Dim tbl As System.Data.DataTable
        Dim int1 As Short
        Dim var1
        Dim strF As String
        Dim strS As String

        dgv = frmH.dgvReports
        tbl = tblReports
        intC = tblReports.Columns.Count

        For Count1 = 0 To intC - 1
            strN = tbl.Columns.Item(Count1).ColumnName
            str1 = "yaya"
            Select Case strN
                Case "CHARREPORTNUMBER"
                    str1 = "CHARREPORTNUMBER"
                    str2 = "Report Number"
                    boolRO = False
                    int1 = 0
                    var1 = DataGridViewContentAlignment.BottomLeft
                    dgv.Columns.Item(str1).MinimumWidth = 75
                Case "CHARREPORTTITLE"
                    str1 = "CHARREPORTTITLE"
                    str2 = "Report Title"
                    boolRO = False
                    dgv.Columns.Item(str1).MinimumWidth = 300
                    int1 = 1
                    var1 = DataGridViewContentAlignment.BottomLeft
                Case "DTREPORTDRAFTISSUEDATE"
                    str1 = "DTREPORTDRAFTISSUEDATE"
                    str2 = "Draft Date"
                    boolRO = False
                    int1 = 2
                    var1 = DataGridViewContentAlignment.BottomCenter
                Case "DTREPORTFINALISSUEDATE"
                    str1 = "DTREPORTFINALISSUEDATE"
                    str2 = "Issue Date"
                    boolRO = False
                    int1 = 3
                    var1 = DataGridViewContentAlignment.BottomCenter
                Case "CHARREPORTTYPE"
                    str1 = "CHARREPORTTYPE"
                    str2 = "Study Type *"
                    boolRO = False
                    int1 = 4
                    var1 = DataGridViewContentAlignment.BottomLeft
                    'Case "CHARREPORTTEMPLATE"
                    '    str1 = "CHARREPORTTEMPLATE"
                    '    str2 = "MSWord Doc Template *"
                    '    boolRO = False
                    '    int1 = 5
                    '    var1 = DataGridViewContentAlignment.BottomLeft
                    '    'Case "ID_TBLCONFIGREPORTTYPE"
                    '    '    str1 = "ID_TBLCONFIGREPORTTYPE"
                    '    '    str2 = "ID_TBLCONFIGREPORTTYPE"
                    '    '    boolRO = True
                    '    '    int1 = 6
                    '    '    var1 = DataGridViewContentAlignment.BottomCenter
            End Select

            If StrComp(strN, str1, CompareMethod.Text) = 0 Then
                'dgv.Columns.item(strN).Visible = True
                'dgv.Columns.item(strN).HeaderText = str2
                'dgv.Columns.item(strN).ReadOnly = boolRO
                dgv.Columns.Item(strN).DisplayIndex = int1
                'dgv.Columns.item(strN).DefaultCellStyle.Alignment = var1
            End If
        Next
    End Sub

    Sub OrderReportsHome()

        frmH.dgvReports.Columns.Item("CHARREPORTNUMBER").DisplayIndex = 0
        frmH.dgvReports.Columns.Item("CHARREPORTTITLE").DisplayIndex = 1
        frmH.dgvReports.Columns.Item("DTREPORTDRAFTISSUEDATE").DisplayIndex = 2
        frmH.dgvReports.Columns.Item("DTREPORTFINALISSUEDATE").DisplayIndex = 3
        frmH.dgvReports.Columns.Item("CHARREPORTTYPE").DisplayIndex = 4

    End Sub

    Sub CreateQCSampleTables()
        Dim Count1 As Short
        Dim str1 As String
        Dim int1 As Short

        'initialize tblQCConcs, used for displaying data
        'for reference: Dim arrBCQCConcs(7, 100) '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=AssayID
        For Count1 = 1 To 8
            Select Case Count1
                Case 1
                    str1 = "LevelNumber"
                Case 2
                    str1 = "Concentration"
                Case 3
                    str1 = "RunID"
                Case 4
                    str1 = "EliminatedFlag"
                Case 5
                    str1 = "SampleName"
                Case 6
                    str1 = "AliquotFactor"
                Case 7
                    str1 = "AssayID"
                Case 8
                    str1 = "AnalyteID"
            End Select
            Dim col As New DataColumn
            col.ColumnName = str1
            tblQCConcs.Columns.Add(col)
        Next

        'initialize tblSampleConcs, used for displaying data
        'for reference: Dim arrBCQCConcs(7, 100) '1=LevelNumber, 2=Concentration, 3=RunID, 4=EliminatedFlag,5=SampleName, 6=AliquotFactor(DilFactor), 7=AssayID
        For Count1 = 1 To 8
            Select Case Count1
                Case 1
                    str1 = "LevelNumber"
                Case 2
                    str1 = "Concentration"
                Case 3
                    str1 = "RunID"
                Case 4
                    str1 = "EliminatedFlag"
                Case 5
                    str1 = "SampleName"
                Case 6
                    str1 = "AliquotFactor"
                Case 7
                    str1 = "AssayID"
                Case 8
                    str1 = "AnalyteID"
            End Select
            Dim col As New DataColumn
            col.ColumnName = str1
            tblSampleConcs.Columns.Add(col)
        Next

    End Sub


    Sub CreatetblAttachment()

        Dim col1 As New DataColumn
        Dim col2 As New DataColumn
        Dim col3 As New DataColumn
        Dim col4 As New DataColumn
        Dim col5 As New DataColumn
        Dim col6 As New DataColumn

        'configure the following table

        col1.ColumnName = "AttachmentNumber"
        col1.DataType = System.Type.GetType("System.Int16")
        tblAttachments.Columns.Add(col1)
        col2.ColumnName = "AnalyteName"
        tblAttachments.Columns.Add(col2)
        col3.ColumnName = "AttachmentName" 'Chromatogram or LM or StudySummary
        tblAttachments.Columns.Add(col3)
        col4.ColumnName = "RepWatsonID" 'For Chrom only, 0 if LM
        col4.DataType = System.Type.GetType("System.Int16")
        tblAttachments.Columns.Add(col4)
        col5.ColumnName = "NumRow" 'row number of tables
        col5.DataType = System.Type.GetType("System.Int16")
        tblAttachments.Columns.Add(col5)
        col6.ColumnName = "CHARFCID"
        tblAttachments.Columns.Add(col6)


    End Sub

    Sub CreatetblAppendix()

        Dim col1 As New DataColumn
        Dim col2 As New DataColumn
        Dim col3 As New DataColumn
        Dim col4 As New DataColumn
        Dim col5 As New DataColumn
        Dim col6 As New DataColumn

        'configure the following table

        col1.ColumnName = "AppendixNumber"
        col1.DataType = System.Type.GetType("System.Int16")
        tblAppendix.Columns.Add(col1)
        col2.ColumnName = "AnalyteName"
        tblAppendix.Columns.Add(col2)
        col3.ColumnName = "AppendixName" 'Chromatogram or LM or StudySummary
        tblAppendix.Columns.Add(col3)
        col4.ColumnName = "RepWatsonID" 'For Chrom only, 0 if LM
        col4.DataType = System.Type.GetType("System.Int16")
        tblAppendix.Columns.Add(col4)
        col5.ColumnName = "NumRow" 'row number of tables
        col5.DataType = System.Type.GetType("System.Int16")
        tblAppendix.Columns.Add(col5)
        col6.ColumnName = "CHARFCID"
        tblAppendix.Columns.Add(col6)


    End Sub

    Sub CreatetblFigures()

        Dim col1 As New DataColumn
        Dim col2 As New DataColumn
        Dim col3 As New DataColumn
        Dim col4 As New DataColumn
        Dim col5 As New DataColumn
        Dim col6 As New DataColumn

        'configure the following table

        Dim tbl As System.Data.DataTable
        tbl = tblFigures

        col1.ColumnName = "FigureNumber"
        col1.DataType = System.Type.GetType("System.Int16")
        tbl.Columns.Add(col1)
        col2.ColumnName = "AnalyteName"
        tbl.Columns.Add(col2)
        col3.ColumnName = "FigureName" 'Chromatogram or LM or StudySummary
        tbl.Columns.Add(col3)
        col4.ColumnName = "RepWatsonID" 'For Chrom only, 0 if LM
        col4.DataType = System.Type.GetType("System.Int16")
        tbl.Columns.Add(col4)
        col5.ColumnName = "NumRow" 'row number of tables
        col5.DataType = System.Type.GetType("System.Int16")
        tbl.Columns.Add(col5)
        col6.ColumnName = "CHARFCID"
        tbl.Columns.Add(col6)


    End Sub

    Sub CreateTableN()

        Dim col1 As New DataColumn
        Dim col2 As New DataColumn
        Dim col3 As New DataColumn
        Dim col4 As New DataColumn
        Dim col5 As New DataColumn
        Dim col6 As New DataColumn

        'configure the following table
        col1.ColumnName = "TableNumber"
        col1.DataType = System.Type.GetType("System.Int16")
        tblTableN.Columns.Add(col1)
        col2.ColumnName = "AnalyteName"
        tblTableN.Columns.Add(col2)
        col3.ColumnName = "TableName"
        tblTableN.Columns.Add(col3)
        col4.ColumnName = "TableID"
        col4.DataType = System.Type.GetType("System.Int16")
        tblTableN.Columns.Add(col4)

        col5.ColumnName = "CHARFCID"
        tblTableN.Columns.Add(col5)

        col6.ColumnName = "TableNameNew"
        tblTableN.Columns.Add(col6)


    End Sub

    Sub SampleReceiptChange()

        'Dim dv as system.data.dataview
        Dim str1 As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim col As DataColumn
        Dim dt As System.Data.DataTable
        Dim strS As String

        dt = tblSampleReceipt
        str1 = "id_tblStudies = " & id_tblStudies
        strS = "dtShipmentReceived ASC"
        Dim dv As System.Data.DataView = New DataView(dt, str1, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowNew = False
        dgv = frmH.dgvSampleReceipt
        dgv.DataSource = dv

        'determine checkbox status
        int1 = dgv.Rows.Count
        If int1 = 0 Then
            frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Unchecked
            frmH.chkManualSampleNumber.CheckState = CheckState.Unchecked
        Else

            If dv(0).Item("boolUseWatson") = 0 Then
                frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Unchecked
            Else
                frmH.chkUseWatsonSampleNumber.CheckState = CheckState.Checked
            End If
            If dv(0).Item("boolUseManual") = 0 Then
                frmH.chkManualSampleNumber.CheckState = CheckState.Unchecked
            Else
                frmH.chkManualSampleNumber.CheckState = CheckState.Checked
            End If

        End If

        'update sample count text boxes
        Call CalcSampleCount()

        'update manual text box, if needed
        If int1 = 0 Then
        Else
            If dv(0).Item("boolUseManual") = 0 Then
            Else
                int1 = dv(0).Item("numTotalSampleNum")
                frmH.txtSRecTotalReport.Text = int1
            End If
        End If

        'populate boolU
        Dim bool As Boolean
        For Count1 = 0 To dv.Count - 1
            dv(Count1).BeginEdit()
            int1 = dv(Count1).Item("boolUse")
            If int1 = -1 Then
                bool = True
            Else
                bool = False
            End If
            dv(Count1).Item("boolU") = bool
            dv(Count1).EndEdit()
        Next

        dgv.AutoResizeColumns()

    End Sub

    Sub SampleReceiptInitialize()
        'Dim dvW as system.data.dataview
        Dim str1 As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim dt As System.Data.DataTable
        Dim var1, var2

        'add unbound column to dt
        If tblSampleReceipt.Columns.Contains("boolU") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "boolU"
            col1.DataType = System.Type.GetType("System.Boolean")
            col1.Caption = "boolU"
            tblSampleReceipt.Columns.Add(col1)
        End If
        dt = tblSampleReceipt

        Dim dv As System.Data.DataView = New DataView(tblSampleReceipt)

        str1 = "id_tblStudies = 0"
        dv.RowFilter = str1

        dgv = frmH.dgvSampleReceipt

        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        dgv.DataSource = dv
        int1 = dgv.Columns.Count
        'configure stuf
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv.Columns.Item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        'configure visible
        dgv.Columns.Item("id_tblSampleReceipt").Visible = False
        dgv.Columns.Item("id_tblStudies").Visible = False
        dgv.Columns.Item("numSampleNumber").Visible = True
        dgv.Columns.Item("numSampleNumber").HeaderText = "Sample Count"
        dgv.Columns.Item("numSampleNumber").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        dgv.Columns.Item("dtShipmentReceived").Visible = True
        dgv.Columns.Item("dtShipmentReceived").HeaderText = "Date Received"
        dgv.Columns.Item("dtShipmentReceived").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("charStorageTemp").Visible = True
        dgv.Columns.Item("charStorageTemp").HeaderText = "Storage Temperature (" & ChrW(176) & "C)"
        dgv.Columns.Item("charStorageTemp").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("charCondition").Visible = True
        dgv.Columns.Item("charCondition").HeaderText = "Sample Condition"
        dgv.Columns.Item("charCondition").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

        dgv.Columns.Item("intOrder").Visible = False

        dgv.Columns.Item("CHARSOURCE").Visible = True
        dgv.Columns.Item("CHARSOURCE").HeaderText = "Source"

        dgv.Columns.Item("boolUse").Visible = False

        dgv.Columns.Item("boolU").Visible = True
        dgv.Columns.Item("boolU").HeaderText = "A_*"
        dgv.Columns.Item("boolU").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("boolUseWatson").Visible = False
        dgv.Columns.Item("boolUseManual").Visible = False
        dgv.Columns.Item("numTotalSampleNum").Visible = False
        dgv.Columns.Item("UPSIZE_TS").Visible = False


        dgv.RowHeadersWidth = 25
        'dgv.ColumnHeadersVisible = True
        dgv.AllowUserToOrderColumns = False
        dgv.AllowUserToResizeColumns = True

        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        'dgv.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect

        Call ReorderSRec()

        

    End Sub

    Sub InitWatsonSampleReciept()


        ''now configure dgvSampleReceiptWatson
        'For Count1 = 1 To 7
        '    Dim col As New DataColumn
        '    Select Case Count1
        '        Case 1
        '            col.ColumnName = "Watson ID"
        '        Case 2
        '            col.ColumnName = "Date Received"
        '        Case 3
        '            col.ColumnName = "Sample Count"
        '        Case 4
        '            col.ColumnName = "Storage Temperature"
        '        Case 5
        '            col.ColumnName = "Sample Condition"
        '        Case 6
        '            col.ColumnName = "STUDYID"
        '        Case 7
        '            col.ColumnName = "Comments"
        '    End Select
        '    tblSRecWatson.Columns.Add(col)
        'Next

        Dim int1 As Int32
        Dim Count1 As Int32

        'configure dgv
        Try
            frmH.dgvSampleReceiptWatson.AllowUserToResizeColumns = True
            frmH.dgvSampleReceiptWatson.AllowUserToResizeRows = True
            frmH.dgvSampleReceiptWatson.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
            frmH.dgvSampleReceiptWatson.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

            'Dim dvW As System.Data.DataView = New DataView(tblSRecWatson)
            ''dvW = frmh.tblSRecWatson.DefaultView
            ''str1 = "STUDYID = 0"
            ''dvW.RowFilter = str1
            'frmH.dgvSampleReceiptWatson.DataSource = dvW

            'str1 = "SELECT DISTINCT Format([DATERECEIVED],""mm/dd/yyyy"") AS DATEREC, SHIPMENT.DATERECEIVED, Count(CONTAINERSAMPLE.DESIGNSAMPLEID) AS SAMPLECOUNT, 
            'STORAGELOCATION.TEMPERATURE, First(SHIPMENT.COMMENTMEMO) AS FirstOfCOMMENTMEMO "

            'configure alignment
            int1 = frmH.dgvSampleReceiptWatson.Columns.Count

            For Count1 = 0 To int1 - 1  '
                frmH.dgvSampleReceiptWatson.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                'frmH.dgvSampleReceiptWatson.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
                frmH.dgvSampleReceiptWatson.Columns.Item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next

            frmH.dgvSampleReceiptWatson.Columns.Item("DATEREC").HeaderText = "Date Received"
            frmH.dgvSampleReceiptWatson.Columns.Item("DATERECEIVED").Visible = False
            frmH.dgvSampleReceiptWatson.Columns.Item("SAMPLECOUNT").HeaderText = "Sample Count"
            frmH.dgvSampleReceiptWatson.Columns.Item("TEMPERATURE").HeaderText = "Storage Temperature (" & ChrW(176) & "C)"
            frmH.dgvSampleReceiptWatson.Columns.Item("FirstOfCOMMENTMEMO").HeaderText = "Sample Condition"

            frmH.dgvSampleReceiptWatson.Columns.Item("FirstOfCOMMENTMEMO").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft


            frmH.dgvSampleReceiptWatson.RowHeadersWidth = 25
            'dgv.ColumnHeadersVisible = True
            frmH.dgvSampleReceiptWatson.AllowUserToOrderColumns = False
            frmH.dgvSampleReceiptWatson.AllowUserToResizeColumns = True

            'frmH.txtSRecTotalReport.Text = 0
            'frmH.txtSRecTotal.Text = 0
            'frmH.txtSRecTotalReportWatson.Text = 0
        Catch ex As Exception

        End Try

    End Sub

    Sub ReorderSRec()

        Dim dv As System.Data.DataView
        Dim str1 As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim col As DataColumn
        Dim dt As System.Data.DataTable
        Dim var1, var2

        dgv = frmH.dgvSampleReceipt

        'set column order
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).DisplayIndex = int1 - 1
        Next

        're-order column
        dgv.Columns.Item("boolU").DisplayIndex = 0
        dgv.Columns.Item("dtShipmentReceived").DisplayIndex = 1
        dgv.Columns.Item("numSampleNumber").DisplayIndex = 2
        dgv.Columns.Item("charStorageTemp").DisplayIndex = 3
        dgv.Columns.Item("charCondition").DisplayIndex = 4

    End Sub

    Sub QATableInitialize()

        Dim dtbl As System.Data.DataTable
        Dim dr1() As DataRow
        Dim ct1 As Short
        Dim tbl1 As System.Data.DataTable
        'Dim tbl2 as System.Data.DataTable
        'Dim tbl3 as System.Data.DataTable
        Dim tbl4 As System.Data.DataTable
        'Dim dr2() As DataRow
        'Dim dr3() As DataRow
        Dim dr4() As DataRow
        'Dim ct2 As Short
        'Dim ct3 As Short
        Dim ct4 As Short
        Dim strF As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim col1 As DataColumn
        Dim dg As DataGrid
        Dim dv As System.Data.DataView
        Dim row As DataRow
        Dim var1, var2, var3
        Dim gs As DataGridColumnStyle
        Dim ctCols As Short
        Dim str1 As String
        Dim int1 As Short

        'set reference to raw data
        tbl1 = tblQATableTemp
        strF = "id_tblConfigReportTables = 1000"
        tbl4 = tblReportTableHeaderConfig
        strF = "id_tblConfigReportTables = 1000 AND id_tblStudies = " & id_tblStudies
        dr4 = tbl4.Select(strF, "intOrder ASC")
        ct4 = dr4.Length
        dg = frmH.dgQATable

        'first configure columns
        If ct4 > 0 Then 'load from existing
            'generate tablestyle
            Dim ts1 As DataGridTableStyle
            ts1 = frmH.dgQATable.TableStyles(0)
            'delete gridstyles 1-n
            Count2 = -1
            For Each gs In ts1.GridColumnStyles
                Count2 = Count2 + 1
            Next
            For Count1 = Count2 To 1 Step -1
                ts1.GridColumnStyles.RemoveAt(Count1)
            Next

            ctCols = 1 'for autosizegrid purposes
            For Count1 = 0 To ct4 - 1
                'evaluate boolInclude
                var1 = dr4(Count1).Item("boolInclude")
                If var1 = -1 Then 'proceed
                    ctCols = ctCols + 1
                    Dim gc1 As New DataGridTextBoxColumn
                    gc1.MappingName = "dtColumn" & Count1 + 1 'dr3(Count1).Item("charColumnLabel")
                    gc1.HeaderText = dr4(Count1).Item("charUserLabel")
                    gc1.NullText = ""
                    'gc1.Width = 75
                    gc1.ReadOnly = False
                    'gc1.Format = "MM/dd/yyyy"
                    gc1.Alignment = HorizontalAlignment.Center
                    ts1.GridColumnStyles.Add(gc1)
                End If
            Next

            'apply ts1 to dg
            ts1.AllowSorting = False
            dg.TableStyles.Clear()
            dg.TableStyles.Add(ts1)
            dv = tbl1.DefaultView
            dv.AllowNew = False
            dv.AllowDelete = False
            dv.AllowEdit = True
            'str1 = "intOrder ASC"
            'dv.Sort = str1
            dg.DataSource = dv
            dg.Refresh()

        Else 'create a new table

            Dim ts1 As New DataGridTableStyle
            'create columns
            If tbl1.Columns.Count = 0 Then

                Dim col10 As New DataColumn
                col10.DataType = System.Type.GetType("System.Int64")
                col10.ColumnName = "ID_QATEMPID" 'for finding deleted rows
                col10.ReadOnly = False
                tbl1.Columns.Add(col10)
                Dim col9 As New DataColumn
                col9.DataType = System.Type.GetType("System.Int64")
                col9.ColumnName = "id_tblReportTableHeaderConfig"
                col9.ReadOnly = False
                tbl1.Columns.Add(col9)
                Dim col8 As New DataColumn
                col8.DataType = System.Type.GetType("System.Int64")
                col8.ColumnName = "id_tblQATables"
                col8.ReadOnly = False
                tbl1.Columns.Add(col8)
                Dim col4 As New DataColumn
                col4.DataType = System.Type.GetType("System.Int64")
                col4.ColumnName = "id_tblStudies"
                col4.ReadOnly = False
                tbl1.Columns.Add(col4)
                Dim col3 As New DataColumn
                col3.DataType = System.Type.GetType("System.Int64")
                col3.ColumnName = "id_tblReports"
                col3.ReadOnly = False
                tbl1.Columns.Add(col3)
                Dim col5 As New DataColumn
                col5.DataType = System.Type.GetType("System.String")
                col5.ColumnName = "charUserLabel"
                col5.ReadOnly = False
                tbl1.Columns.Add(col5)
                Dim col6 As New DataColumn
                col6.DataType = System.Type.GetType("System.Int16")
                col6.ColumnName = "intOrder"
                col6.ReadOnly = False
                tbl1.Columns.Add(col6)
                'add the next column for column0 data in grid
                'Dim col7 As New DataColumn
                'col7.DataType = System.Type.GetType("System.String")
                'col7.ColumnName = "charCriticalPhase"
                'col7.ReadOnly = False
                'tbl1.Columns.Add(col7)

                For Count1 = 1 To 8
                    Dim col2 As New DataColumn
                    'col2.DataType = System.Type.GetType("System.DateTime")
                    col2.DataType = System.Type.GetType("System.String")
                    col2.ColumnName = "dtColumn" & Count1
                    col2.ReadOnly = False
                    tbl1.Columns.Add(col2)
                Next

            Else

            End If

            'configure tablestyle
            Dim gc As New DataGridTextBoxColumn
            gc.MappingName = "charUserLabel"
            gc.HeaderText = "Critical Phase"
            gc.NullText = ""
            gc.Width = 175
            gc.ReadOnly = True
            ts1.GridColumnStyles.Add(gc)
            For Count1 = 1 To 8
                Dim gc1 As New DataGridTextBoxColumn
                gc1.MappingName = "dtColumn" & Count1
                gc1.HeaderText = "Column " & Count1
                gc1.NullText = ""
                'gc1.Width = 75
                'gc1.Format = "MM/dd/yyyy"
                gc1.ReadOnly = False
                gc1.Alignment = HorizontalAlignment.Center
                ts1.GridColumnStyles.Add(gc1)
            Next
            'Dim gc2 As New DataGridTextBoxColumn
            'gc2.MappingName = "intOrder"
            'gc2.HeaderText = "Order"
            'gc2.NullText = ""
            'gc2.Width = 75
            'gc2.ReadOnly = True
            'ts1.GridColumnStyles.Add(gc2)

            ctCols = 9

            dv = tbl1.DefaultView
            dv.AllowNew = False
            dv.AllowEdit = True
            dv.AllowDelete = False
            ts1.AllowSorting = False
            dg.TableStyles.Clear()
            dg.TableStyles.Add(ts1)
            dg.DataSource = dv
            dg.Refresh()

        End If

        'now create rows
        dtbl = tblQATables
        strF = "id_tblStudies = " & id_tblStudies
        dr1 = dtbl.Select(strF, "intOrder ASC")
        ct1 = dr1.Length
        'tbl4 = tblReportTableHeaderConfig
        strF = "id_tblConfigReportTables = 2000 AND id_tblStudies = " & id_tblStudies & " AND boolInclude = -1"
        dr4 = tbl4.Select(strF, "intOrder ASC")
        ct4 = dr4.Length

        tbl1.Rows.Clear()
        'check to see if data exists in dtbl
        If ct1 = 0 Then 'data doesn't exist. Use default from tbl4
            'For Count1 = 0 To ct4 - 1
            '    row = tbl1.NewRow
            '    var1 = dr4(Count1).Item("id_tblReportTableHeaderConfig")
            '    row("id_tblReportTableHeaderConfig") = var1
            '    row("id_tblQATables") = 0
            '    row("id_tblStudies") = id_tblStudies
            '    row("id_tblReports") = 0
            '    var1 = dr4(Count1).Item("charUserLabel")
            '    row("charUserLabel") = var1
            '    var1 = dr4(Count1).Item("intOrder")
            '    row("intOrder") = var1
            '    tbl1.Rows.Add(row)
            'Next
        Else 'data exists. Use dtbl
            For Count1 = 0 To ct1 - 1
                row = tbl1.NewRow
                For Each col1 In dtbl.Columns
                    If StrComp(col1.ColumnName, "UPSIZE_TS", CompareMethod.Text) = 0 Then
                    Else
                        row(col1.ColumnName) = dr1(Count1).Item(col1.ColumnName)
                    End If
                Next
                'find charUserLabel
                var1 = row("id_tblReportTableHeaderConfig")
                'strF = "id_tblConfigReportTables = 2000 AND id_tblStudies = " & id_tblStudies & " AND boolInclude = -1 AND id_tblReportTableHeaderConfig = " & var1
                strF = "id_tblConfigReportTables = 2000 AND id_tblStudies = " & id_tblStudies & "  AND id_tblReportTableHeaderConfig = " & var1
                dr4 = tbl4.Select(strF)
                int1 = dr4.Length
                If int1 = 0 Then
                    row("charUserLabel") = ""
                Else
                    row("charUserLabel") = dr4(0).Item("charUserLabel")
                End If
                tbl1.Rows.Add(row)
            Next

            'now accept changes to tbl1 so that Audit Trail code works
            tbl1.AcceptChanges()

        End If

        Call AutoSizeGrid(175, dv, dg, dv.Count, ctCols, 0, False)

    End Sub

    Sub DoCancelQATable()
        Call QATableInitialize()
    End Sub

    Sub SaveSampleReceiptTab()

        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim var1
        Dim strM As String

        Try
            dv = frmH.dgvSampleReceipt.DataSource
            int2 = CInt(frmH.txtSRecTotalReport.Text)
            int1 = dv.Count

            'finalize numTotalSampleNum
            For Count1 = 0 To int1 - 1
                dv(Count1).BeginEdit()
                dv(Count1).Item("numTotalSampleNum") = int2
                dv(Count1).EndEdit()
            Next

            Dim dvCheck As System.Data.DataView = New DataView(tblSampleReceipt)
            dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
            Dim int10 As Short
            int10 = 1
            If int10 = 0 Then
            Else

                Call FillAuditTrailTemp(tblSampleReceipt)

                If boolGuWuOracle Then
                    Try
                        ta_tblSampleReceipt.Update(tblSampleReceipt)
                    Catch ex As DBConcurrencyException
                        'ds2005.TBLSAMPLERECEIPT.Merge('ds2005.TBLSAMPLERECEIPT, True)
                    End Try
                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblSampleReceiptAcc.Update(tblSampleReceipt)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLSAMPLERECEIPT.Merge('ds2005Acc.TBLSAMPLERECEIPT, True)
                    End Try
                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblSampleReceiptSQLServer.Update(tblSampleReceipt)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLSAMPLERECEIPT.Merge('ds2005Acc.TBLSAMPLERECEIPT, True)
                    End Try
                End If

            End If
            frmH.dgvSampleReceipt.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        Catch ex As Exception
            var1 = ex.Message
            strM = "Problem executing 'SaveSampleReceiptTab'" & ChrW(10) & ChrW(10) & var1
            strM = strM & ChrW(10) & ChrW(10) & "StudyDoc exeuction will continue."
            MsgBox(strM, vbInformation, "Problem...")
            var1 = var1
        End Try


    End Sub

    Sub SaveQATable()

        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim tbl3 As System.Data.DataTable
        Dim tbl4 As System.Data.DataTable
        Dim drows1() As DataRow
        Dim drows2() As DataRow
        Dim drows3() As DataRow
        Dim drows4() As DataRow
        Dim drowsF() As DataRow
        Dim ct1 As Short
        Dim ct2 As Short
        Dim ct3 As Short
        Dim ct4 As Short
        Dim ctF As Short
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2, var3, var4, var5
        Dim str1 As String
        Dim strF As String
        Dim int1 As Short
        Dim int2 As Short
        Dim intS As Long
        Dim strS As String
        Dim boolExists As Boolean
        Dim row As DataRow
        Dim rowsT() As DataRow
        Dim drowsMaxID() As DataRow
        Dim maxID
        Dim maxID1
        Dim arrMaxID(100) As Int64

        'find maxID for tblReportTable
        maxID = GetMaxID("tblQATables", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid

        'str1 = "charTable = 'tblQATables'"
        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If
        'drowsMaxID = tblMaxID.Select(str1)
        'maxID = drowsMaxID(0).Item("numMaxID")
        maxID1 = maxID

        tbl1 = tblQATables
        tbl2 = tblQATableTemp
        ct2 = tbl2.Rows.Count
        tbl3 = tblConfigHeaderLookup
        ct3 = tbl3.Rows.Count

        'start Audit Trail approach

        Dim intTotal As Short = tbl2.Rows.Count 'debug
        Dim dvQA As System.Data.DataView = New DataView(tbl2)
        'dvQA.RowStateFilter = DataViewRowState.ModifiedOriginal
        'Dim intMO As Short = dvQA.Count
        dvQA.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim intMC As Short = dvQA.Count

        'enter modified information
        For Count1 = 1 To intMC
            'var1 = tbl2.Rows.item(Count1).Item("id_tblQATables") 'from tblQATableTemp
            var1 = dvQA(Count1 - 1).Item("id_tblStudies") 'from tblQATableTemp
            var2 = dvQA(Count1 - 1).Item("id_tblReportTableHeaderConfig") 'from tblQATableTemp id_tblReportTableHeaderConfig
            strF = "id_tblStudies = " & var1 & " AND id_tblReportTableHeaderConfig = " & var2

            var1 = dvQA(Count1 - 1).Item("ID_TBLQATABLES")
            strF = "ID_TBLQATABLES = " & var1

            Erase rowsT
            rowsT = tbl1.Select(strF)
            'If rowsT.Length = 0 Then
            '    boolExists = False
            'Else
            '    boolExists = True
            'End If
            'If boolExists Then 'update row
            '    'row = tbl1.Rows.item(Count1)
            '    row = rowsT(0) 'drows1(Count2)
            '    row.BeginEdit()
            'Else 'add a new row to tbl1
            '    row = tbl1.NewRow
            '    row.BeginEdit()
            '    maxID = maxID + 1
            '    row("id_tblQATables") = maxID
            'End If
            row = rowsT(0) 'drows1(Count2)
            row.BeginEdit()

            row("id_tblStudies") = id_tblStudies
            row("id_tblReports") = dvQA(Count1 - 1).Item("id_tblReports")
            row("id_tblReportTableHeaderConfig") = dvQA(Count1 - 1).Item("id_tblReportTableHeaderConfig")
            ct3 = dvQA(Count1 - 1).Item("intOrder") 'for debugging
            row("intOrder") = dvQA(Count1 - 1).Item("intOrder")
            For Count2 = 1 To 8
                row("dtColumn" & Count2) = dvQA(Count1 - 1).Item("dtColumn" & Count2)
            Next
            'If boolExists Then 'update row
            '    row.EndEdit()
            'Else 'add row
            '    row.EndEdit()
            '    tbl1.Rows.Add(row)
            'End If
            row.EndEdit()
        Next

        dvQA.RowStateFilter = DataViewRowState.Added
        Dim intAdded As Short = dvQA.Count
        'now do added rows
        For Count1 = 1 To intAdded
            maxID = maxID + 1
            row = tbl1.NewRow
            row.BeginEdit()
            row("id_tblQATables") = maxID
            row("id_tblStudies") = id_tblStudies
            row("id_tblReports") = 0
            row("id_tblReportTableHeaderConfig") = dvQA(Count1 - 1).Item("id_tblReportTableHeaderConfig")
            'row("charUserLabel") = tbl2.Rows.item(Count1).Item("charUserLabel")
            ct3 = dvQA(Count1 - 1).Item("intOrder") 'for debugging
            row("intOrder") = dvQA(Count1 - 1).Item("intOrder")
            For Count2 = 1 To 8
                row("dtColumn" & Count2) = dvQA(Count1 - 1).Item("dtColumn" & Count2)
            Next
            row.EndEdit()
            tbl1.Rows.Add(row)
        Next

        dvQA.RowStateFilter = DataViewRowState.Deleted
        Dim intDeleted As Short = dvQA.Count
        'now do deleted rows
        For Count1 = 1 To intDeleted
            var1 = NZ(dvQA(Count1 - 1).Item("ID_QATEMPID"), "") 'from tblQATableTemp
            var2 = NZ(dvQA(Count1 - 1).Item("ID_TBLQATABLES"), "") 'from tblQATableTemp

            If Len(var2) = 0 Then
                strF = "ID_QATEMPID = " & var1
            ElseIf Len(var1) = 0 Then
                strF = "ID_TBLQATABLES = " & var2
            End If
            Erase rowsT
            rowsT = tbl1.Select(strF)
            For Count2 = rowsT.Length - 1 To 0 Step -1
                rowsT(Count2).Delete()
            Next

        Next

        GoTo skip1

        'end Audit Trail approach

        ''first see if there are existing data
        'strF = "id_tblStudies = " & id_tblStudies
        'drows1 = tbl1.Select(strF)
        'ct1 = drows1.Length
        If ct1 = 0 Then 'add brand new records
            For Count1 = 0 To ct2 - 1
                maxID = maxID + 1
                row = tbl1.NewRow
                row("id_tblQATables") = maxID
                row("id_tblStudies") = id_tblStudies
                row("id_tblReports") = 0
                row("id_tblReportTableHeaderConfig") = tbl2.Rows.Item(Count1).Item("id_tblReportTableHeaderConfig")
                'row("charUserLabel") = tbl2.Rows.item(Count1).Item("charUserLabel")
                ct3 = tbl2.Rows.Item(Count1).Item("intOrder") 'for debugging
                row("intOrder") = tbl2.Rows.Item(Count1).Item("intOrder")
                For Count2 = 1 To 8
                    row("dtColumn" & Count2) = tbl2.Rows.Item(Count1).Item("dtColumn" & Count2)
                Next
                tbl1.Rows.Add(row)
            Next
        ElseIf ct1 > 0 Then 'update and add/delete rows as needed
            For Count1 = 0 To ct2 - 1
                'var1 = tbl2.Rows.item(Count1).Item("id_tblQATables") 'from tblQATableTemp
                var1 = tbl2.Rows.Item(Count1).Item("id_tblStudies") 'from tblQATableTemp
                var2 = tbl2.Rows.Item(Count1).Item("id_tblReportTableHeaderConfig") 'from tblQATableTemp
                strF = "id_tblStudies = " & var1 & " AND id_tblReportTableHeaderConfig = " & var2
                Erase rowsT
                rowsT = tbl1.Select(strF)
                If rowsT.Length = 0 Then
                    boolExists = False
                Else
                    boolExists = True
                End If
                If boolExists Then 'update row
                    'row = tbl1.Rows.item(Count1)
                    row = rowsT(0) 'drows1(Count2)
                    row.BeginEdit()
                Else 'add a new row to tbl1
                    row = tbl1.NewRow
                    row.BeginEdit()
                    maxID = maxID + 1
                    row("id_tblQATables") = maxID
                End If
                row("id_tblStudies") = id_tblStudies
                row("id_tblReports") = tbl2.Rows.Item(Count1).Item("id_tblReports")
                row("id_tblReportTableHeaderConfig") = tbl2.Rows.Item(Count1).Item("id_tblReportTableHeaderConfig")
                ct3 = tbl2.Rows.Item(Count1).Item("intOrder") 'for debugging
                row("intOrder") = tbl2.Rows.Item(Count1).Item("intOrder")
                For Count2 = 1 To 8
                    row("dtColumn" & Count2) = tbl2.Rows.Item(Count1).Item("dtColumn" & Count2)
                Next
                If boolExists Then 'update row
                    row.EndEdit()
                Else 'add row
                    row.EndEdit()
                    tbl1.Rows.Add(row)
                End If
            Next
            ''now check to see if any rows need to be deleted
            'For Count1 = 0 To ct1 - 1
            '    'var1 = drows1(Count1).Item("id_tblQATables") 'from tblQATables
            '    var1 = drows1(Count1).Item("id_tblStudies") 'from tblQATables
            '    var2 = drows1(Count1).Item("id_tblqatables") 'from tblQATables
            '    strF = "id_tblStudies = " & var1 & " AND id_tblqatables = " & var2
            '    Erase rowsT
            '    rowsT = tbl2.Select(strF)
            '    If rowsT.Length = 0 Then
            '        boolExists = False
            '    Else
            '        boolExists = True
            '    End If

            '    'boolExists = False
            '    'For Count2 = 0 To ct2 - 1
            '    '    var2 = tbl2.Rows.item(Count2).Item("id_tblQATables")
            '    '    If var1 = var2 Then 'exists
            '    '        boolExists = True
            '    '        Exit For
            '    '    End If
            '    'Next
            '    If boolExists Then 'ignore
            '    Else 'delete row from tbl1
            '        drows1(Count1).Delete()
            '    End If

            'Next
        Else
        End If

skip1:

        If maxID = maxID1 Then 'now rows were added
        Else

            Call PutMaxID("tblQATables", maxID)

            'drowsMaxID(0).BeginEdit()
            'drowsMaxID(0).Item("numMaxID") = maxID
            'drowsMaxID(0).EndEdit()

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

        frmH.dgQATable.Update()

        Dim dvCheck As System.Data.DataView = New DataView(tblQATables)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblQATables)

            If boolGuWuOracle Then
                Try
                    ta_tblQATables.Update(tblQATables)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLQATABLES.Merge('ds2005.TBLQATABLES, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblQATablesAcc.Update(tblQATables)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLQATABLES.Merge('ds2005Acc.TBLQATABLES, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblQATablesSQLServer.Update(tblQATables)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLQATABLES.Merge('ds2005Acc.TBLQATABLES, True)
                End Try
            End If


        End If


    End Sub

    Sub UpdateWB_RBS()

        If boolFormLoad Then
            Exit Sub
        End If

        If frmH.panRBSwb.Visible Then 'continue
        Else
            Try

                frmH.wbRBS.Navigate("about:blank")

            Catch ex As Exception
                Dim var10
                var10 = "Done'" 'for debugging
            End Try

            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim strPath As String
        Dim id As Int64

        dgv = frmH.dgvReportStatementWord
        If dgv.Rows.Count = 0 Then
            Try
                'frmH.afrRBS.Close()
            Catch ex As Exception

            End Try
            frmH.wbRBS.Navigate("about:blank")
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

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

        dtbl1 = tblWordStatements
        dtbl2 = tblWorddocs

        strPath = NZ(dgv("CHARWORDSTATEMENT", intRow).Value, "") 'don't need
        id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value
        strF = "ID_TBLWORDSTATEMENTS = " & id
        strS = "ID_TBLWORDDOCS ASC"
        rows2 = dtbl2.Select(strF, strS)
        intL = rows2.Length

        'If intL = 0 Then
        '    frmH.wbRBS.Navigate("about:blank")
        '    GoTo end1
        'End If

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        strpathT = ""
        Do Until boolE = False
            Count1 = Count1 + 1
            strpathT = "C:\Labintegrity\StudyDoc\Temp\Temp" & Format(Count1, "00000") & ".xml"
            If File.Exists(strpathT) Then
            Else
                Exit Do
            End If
        Loop

        Dim strBuild = New StringBuilder("")
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

        frmH.wbRBS.Navigate(strpathT)

        'clean up temp file
        'File.Delete(strpathT)

end1:

    End Sub

    Sub ExpandRBS(ByVal boolAFR As Boolean)

        Exit Sub

        Dim h1, h2, w1, w2, t1, t2
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView

        dgv1 = frmH.dgvReportStatements
        dgv2 = frmH.dgvReportStatementWord

        h1 = frmH.tp5.Height
        w1 = frmH.tp5.Width

        If boolAFR Then 'show afr

            If frmH.panRBSwb.Visible Then
            Else
                Dim var1, var2, var3, var4
                var1 = dgv1.Top
                var2 = dgv1.Height
                var3 = dgv2.Top
                var4 = dgv2.Height

                dgv1.Height = dgv1.Height / 2
                dgv2.Height = (dgv1.Height + dgv1.Top) - dgv2.Top 'dgv2.Height / 2

                If frmH.panRBSwb.Height = dgv1.Height Then
                Else

                    frmH.panRBSwb.Top = dgv1.Top + dgv1.Height + 2
                    frmH.panRBSwb.Height = dgv1.Height
                    frmH.panRBSwb.Left = dgv1.Left
                    frmH.panRBSwb.Width = dgv2.Left + dgv2.Width

                End If


                frmH.panRBSwb.Visible = True
            End If



        Else
            dgv1.Height = dgv1.Height * 2
            dgv2.Height = (dgv1.Height + dgv1.Top) - dgv2.Top
            frmH.panRBSwb.Visible = False
            frmH.wbRBS.Navigate("about:blank")

        End If



    End Sub


    Sub ReportStatementGetStatementTitlesFromWord(ByVal frm As Form)

        Dim var1, var2, var3, var4, var5
        Dim strPath As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim dv As System.Data.DataView
        Dim tbl2 As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim dtblPaths As System.Data.DataTable
        Dim row1 As DataRow
        Dim ct As Short
        Dim bool1 As Boolean
        Dim intWdT As Short
        Dim intWdTCt As Short
        Dim strSource As String
        Dim wdTRows As Short
        Dim boolFE As Boolean

        Dim strM As String
        Dim strM1 As String
        Dim ctE As Short
        Dim dgv As DataGridView

        Dim strF As String = "CHARWORDSTATEMENT = 'Active'"

        Try

            dtbl = tblWordStatements

            dgv = frmH.dgvReportStatementWord
            dgv.RowHeadersWidth = 25
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Dim dv10 As System.Data.DataView = New DataView(dtbl, strF, "", DataViewRowState.CurrentRows)
            dv10.AllowNew = False
            dv10.AllowDelete = False
            dv10.AllowEdit = False

            dgv.ReadOnly = True

            dgv.DataSource = dv10

            int1 = dgv.Columns.Count
            For Count1 = 0 To int1 - 1
                dgv.Columns(Count1).Visible = False
                dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            dgv.Columns("CHARTITLE").Visible = True
            dgv.Columns("CHARTITLE").HeaderText = "Statement"
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try




    End Sub

    Sub ReportStatementChangeCbxFill()

        Exit Sub

        Dim dv As System.Data.DataView
        Dim var1, var2, var3, var4, var5
        Dim str1 As String
        Dim tbl2 As System.Data.DataTable
        Dim int1 As Integer
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim Count1 As Short


        dgv = frmH.dgvReportStatements
        'tbl2 = tblReportStatementsStore
        If frmH.dgvReportStatements.RowCount = 0 Then
            Exit Sub
        ElseIf frmH.dgvReportStatements.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = frmH.dgvReportStatements.CurrentRow.Index
        End If
        dv = frmH.dgvReportStatements.DataSource
        'var1 = dv.Item(intRow).Item("id_tblConfigBodySections")
        'var2 = dv.Item(intRow).Item("boolUseStatements")
        'var3 = dv.Item(intRow).Item("boolGuWu")

        var1 = dgv.Rows.Item(intRow).Cells("id_tblConfigBodySections").Value
        var2 = dgv.Rows.Item(intRow).Cells("boolUseStatements").Value
        var3 = dgv.Rows.Item(intRow).Cells("boolGuWu").Value

        'var4 = ""
        'If var2 = -1 Then
        '    var4 = "client"
        'ElseIf var3 = -1 Then
        '    var4 = "guwu"
        'End If

        var4 = "StudyDoc"


        'str1 = "id_tblConfigBodySections = " & var1 & " AND charSource = '" & var4 & "'"
        str1 = "id_tblConfigBodySections = " & var1 & " AND charSource = '" & var4 & "' AND CHARWORDSTATEMENT = 'Active'"
        Dim dv1 As System.Data.DataView
        dv1 = frmH.dgvReportStatementWord.DataSource
        Try
            dv1.RowFilter = str1
        Catch ex As Exception

        End Try

        dgv.AutoResizeRows()

        frmH.dgvReportStatementWord.Refresh()


    End Sub

    Sub SaveReportStatements()

        frmH.dgvReportStatements.CommitEdit(DataGridViewDataErrorContexts.Commit)
        Dim dvCheck As System.Data.DataView = New DataView(tblReportstatements)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblReportstatements)

            If boolGuWuOracle Then
                Try
                    ta_tblReportStatements.Update(tblReportStatements)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLREPORTSTATEMENTS.Merge('ds2005.TBLREPORTSTATEMENTS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportStatementsAcc.Update(tblReportStatements)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTSTATEMENTS.Merge('ds2005Acc.TBLREPORTSTATEMENTS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportStatementsSQLServer.Update(tblReportStatements)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTSTATEMENTS.Merge('ds2005Acc.TBLREPORTSTATEMENTS, True)
                End Try
            End If

        End If

        Call FillAuditTrailTemp(tblTableProperties)

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


        'This gets done in Report Template Window
        'If boolGuWuOracle Then
        '    Try
        '        frmH.ta_tblWordStatements.Update(tblWordStatements)
        '    Catch ex As DBConcurrencyException
        '        'ds2005.TBLWORDSTATEMENTS.Merge('ds2005.TBLWORDSTATEMENTS, True)
        '    End Try
        'ElseIf boolGuWuAccess Or boolGuWuSQLServer Then
        '    Try
        '        frmH.ta_tblWordStatementsAcc.Update(tblWordStatements)
        '    Catch ex As DBConcurrencyException
        '        'ds2005Acc.TBLWORDSTATEMENTS.Merge('ds2005Acc.TBLWORDSTATEMENTS, True)
        '    End Try
        'End If


        'pesky
        Call ReportStatementsFillCharSection()


    End Sub

    Sub ReportStatmentInitialize()

        Dim var1, var2, var3, var4
        Dim strPath As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim dv As System.Data.DataView
        'Dim tbl As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim drows2() As DataRow
        Dim intRows2 As Short
        Dim drows() As DataRow
        Dim col1 As DataColumn
        Dim row1 As DataRow
        'Dim intCols As Short
        Dim intRows As Short
        Dim dg As DataGrid
        Dim col As DataColumn
        Dim boolColName As Boolean
        Dim st As System.Type
        Dim strExpr As String
        Dim arrStr(3) As String
        Dim dgv As DataGridView


        'tbl = tblReportStatementsGuWu
        'intCols = tbl.Columns.Count

        dtbl = tblReportstatements

        str1 = "* Include in Report Body"
        str1 = str1 & ChrW(10) & "* A = Style Heading Level"
        str1 = str1 & ChrW(10) & "* B = Check = Insert Page Break Before Heading Level"

        frmH.lblRBS.Text = str1
        Dim a, b, c
        'a = frmH.dgvReportStatements.Top
        'b = frmH.lblRBS.Height
        'c = a - b
        'frmH.lblRBS.Top = c

        'add columns to tbl
        'For Each col In dtbl.Columns
        '    Dim col3 As New DataColumn
        '    col3.DataType = col.DataType 'System.Type.GetType("System.Int64")
        '    str1 = col.ColumnName
        '    col3.ColumnName = col.ColumnName
        '    col3.ReadOnly = False
        '    tbl.Columns.Add(col3)
        'Next

        Try
            'add unbound columns to dtbl
            Dim col2 As New DataColumn
            col2.DataType = System.Type.GetType("System.String")
            col2.ColumnName = "charSectionName"
            col2.Caption = "Section Name"
            'col2.ReadOnly = True
            dtbl.Columns.Add(col2)

            Dim col3 As New DataColumn
            col3.DataType = System.Type.GetType("System.Boolean")
            col3.ColumnName = "boolI"
            col3.Caption = "Include*"
            'col3.ReadOnly = True
            dtbl.Columns.Add(col3)

            Dim col4 As New DataColumn
            col4.DataType = System.Type.GetType("System.Boolean")
            col4.ColumnName = "boolPB"
            col4.Caption = "B*"
            'col4.ReadOnly = True
            dtbl.Columns.Add(col4)

            'Dim col5 As New DataColumn
            'col5.DataType = System.Type.GetType("System.Boolean")
            'col5.ColumnName = "boolGW"
            'col5.Caption = "C_*"
            ''col5.ReadOnly = True
            'dtbl.Columns.Add(col5)

            'Dim col6 As New DataColumn
            'col6.DataType = System.Type.GetType("System.Int16")
            'col6.ColumnName = "NUMCOMPANY"
            'col6.Caption = "COMPANYID"
            ''col6.ReadOnly = True
            'dtbl.Columns.Add(col6)
        Catch ex As Exception

        End Try


        Dim col10 As DataColumn
        For Each col10 In dtbl.Columns
            Select Case col10.ColumnName
                Case "NUMHEADINGLEVEL"
                    col10.Caption = "A*"
                Case "BOOLINCLUDE"
                    col10.Caption = "Include*"
                Case "INTORDER"
                    col10.Caption = "Order"
                Case "BOOLUSESTATEMENTS"
                    col10.Caption = "B_*"
                Case "BOOLGUWU"
                    col10.Caption = "C_*"
                Case "CHARSTATEMENT"
                    col10.Caption = "Assigned Statement"
                Case "CHARHEADINGTEXT"
                    col10.Caption = "Heading Text"
                    'col10.ReadOnly = True
            End Select
        Next


        frmH.dgvReportStatements.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        frmH.dgvReportStatements.AllowUserToResizeColumns = True
        frmH.dgvReportStatements.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        frmH.dgvReportStatements.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.ColumnHeader)
        frmH.dgvReportStatements.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        frmH.dgvReportStatements.AllowUserToResizeRows = True
        frmH.dgvReportStatements.RowHeadersWidth = 25
        frmH.dgvReportStatements.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        frmH.dgvReportStatements.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect

        'do stuff with dgvreportstatementword
        dgv = frmH.dgvReportStatementWord

        dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.RowHeadersWidth = 25
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders

    End Sub

    Sub ReportStatementsFill()

        Dim var1, var2, var3, var4
        Dim strPath As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim tbl As System.Data.DataTable
        'Dim tbl1 As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim drows() As DataRow
        Dim drows2() As DataRow
        Dim drows3() As DataRow
        Dim intRows2 As Short
        Dim col1 As DataColumn
        Dim row1 As DataRow
        Dim intCols As Short
        Dim intRows As Short
        Dim ts1 As New DataGridTableStyle
        Dim dg As DataGrid
        Dim strF As String
        Dim strS As String
        Dim ct1 As Short
        Dim dv1 As System.Data.DataView
        Dim dgv As DataGridView
        Dim rows() As DataRow
        Dim num1
        Dim tbl3 As System.Data.DataTable
        Dim strF3 As String
        Dim rows3() As DataRow
        Dim intRows3 As Short
        Dim omax, oval
        Dim strM As String
        Dim int4 As Short
        Dim idWS As Int64
        Dim dt As Date
        Dim intWordTableNumber As Int16

        dt = Now

        dgv = frmH.dgvReportStatements
        'tbl = frmh.qryConfigBodySections
        'tbl1 = tblReportStatementsGuWu
        dtbl = tblReportstatements
        'tbl3 = tblReportStatementsStore
        tbl3 = tblWordStatements

        'dtbl.Columns.item("charSectionName").ReadOnly = False
        'dtbl.Columns.item("charStatement").ReadOnly = False

        dtbl.RejectChanges() 'first reject changes because this routine will get called in a cancel action


        '*****
        tbl = tblConfiguration
        strF = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Company ID'"
        rows = tbl.Select(strF)
        num1 = rows(0).Item("CHARCONFIGVALUE")

        strF = "id_tblStudies = " & id_tblStudies

        drows = dtbl.Select(strF)
        intRows = drows.Length

        tbl = tblConfigBodySections
        ct1 = tbl.Rows.Count

        'If boolFormLoad Then
        '    frmH.pb1.Value = oval
        '    frmH.pb1.Maximum = omax
        '    frmH.pb1.Value = 0
        '    frmH.pb1.Maximum = ct1
        'End If

        If intRows = 0 Then 'fill with default values
            boolCont = False
            frmH.rbShowAllRBody.Checked = True
            boolCont = True

            For Count1 = 0 To ct1 - 1

                'If boolFormLoad Then
                '    str1 = "Formatting Report Body grid: " & Count1 & " of " & ct1
                '    frmH.pb1.Value = Count1
                '    frmH.pb1.Refresh()
                '    frmH.lblProgress.Text = str1
                '    frmH.lblProgress.Refresh()
                'End If

                'Dim newDRV As DataRowView = dv.AddNew()
                Dim newDRV As DataRow = dtbl.NewRow
                newDRV("id_tblStudies") = id_tblStudies
                newDRV("id_tblConfigReportType") = tbl.Rows.Item(Count1).Item("id_tblConfigReportType") 'id_tblConfigReportType
                var1 = tbl.Rows.Item(Count1).Item("id_tblConfigBodySections")
                intWordTableNumber = tbl.Rows.Item(Count1).Item("intWordTableNumber")
                newDRV("id_tblConfigBodySections") = var1
                'find 1st charStatement for guwu
                strF3 = "id_tblConfigBodySections = " & var1 ' & " AND charSource = 'guwu'"
                Erase rows3
                intRows = tbl3.Columns.Count
                rows3 = tbl3.Select(strF3)
                intRows3 = rows3.Length
                If intRows3 = 0 Then
                    'must create an entry for tblWordStatements
                    'ID_TBLWORDSTATEMENTS
                    'ID_TBLCONFIGBODYSECTIONS
                    'INTWORDTABLENUMBER
                    'CHARTITLE
                    'CHARWORDSTATEMENT
                    'UPSIZE_TS

                   

                    Dim rowsM() As DataRow
                    Dim strFM As String
                    Dim maxID As Int64
                    maxID = GetMaxID("TBLWORDSTATEMENTS", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
                    '20190219 LEE: Don't need anymore. Used GetMaxID
                    'Call PutMaxID("TBLWORDSTATEMENTS", maxID)

                    'strFM = "CHARTABLE = 'TBLWORDSTATEMENTS'"
                    'rowsM = tblMaxID.Select(strFM)
                    'maxID = rowsM(0).Item("NUMMAXID")
                    'maxID = maxID + 1
                    'rowsM(0).BeginEdit()
                    'rowsM(0).Item("NUMMAXID") = maxID
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


                    Dim newWS As DataRow = tbl3.NewRow
                    newWS.BeginEdit()
                    newWS("ID_TBLWORDSTATEMENTS") = maxID
                    newWS("ID_TBLCONFIGBODYSECTIONS") = var1
                    newWS("INTWORDTABLENUMBER") = intWordTableNumber
                    newWS("CHARTITLE") = "Statement 1"
                    newWS("CHARWORDSTATEMENT") = "Blank" 'this is supposed to be a path. It may be ignored
                    newWS("UPSIZE_TS") = dt
                    newWS.EndEdit()
                    tbl3.Rows.Add(newWS)
                    'Try
                    '    frmH.ta_tblWordStatements.Update(tblWordStatements)
                    'Catch ex As DBConcurrencyException
                    '    'ds2005.TBLWORDSTATEMENTS.Merge('ds2005.TBLWORDSTATEMENTS, True)
                    'End Try

                    var1 = "Statement 1"
                    idWS = maxID
                Else
                    var1 = rows3(0).Item("CHARTITLE")
                    idWS = rows3(0).Item("ID_TBLWORDSTATEMENTS")
                End If
                newDRV("charStatement") = var1
                newDRV("ID_TBLWORDSTATEMENTS") = idWS

                var1 = tbl.Rows.Item(Count1).Item("charSectionName") 'unbound column
                newDRV("charSectionName") = var1 'unbound column

                newDRV("intOrder") = tbl.Rows.Item(Count1).Item("intOrder")
                newDRV("boolInclude") = 0 'false
                newDRV("boolUseStatements") = 0 'False
                newDRV("boolGuWu") = -1 'True
                newDRV("NUMHEADINGLEVEL") = tbl.Rows.Item(Count1).Item("NUMHEADINGLEVEL")
                newDRV("NUMCOMPANY") = tbl.Rows.Item(Count1).Item("NUMCOMPANY")
                'var1 = tbl.Rows.Item(Count1).Item("CHARHEADINGTEXT")
                newDRV("CHARHEADINGTEXT") = tbl.Rows.Item(Count1).Item("CHARHEADINGTEXT")
                newDRV("BOOLPAGEBREAK") = 0 'false
                dtbl.Rows.Add(newDRV)
            Next
            var1 = var1 'debug
        Else
            'update 
            For Count1 = 0 To ct1 - 1 'intRows - 1

                'If boolFormLoad Then
                '    str1 = "Formatting Report Body grid: " & Count1 & " of " & ct1
                '    frmH.pb1.Value = Count1
                '    frmH.pb1.Refresh()
                '    frmH.lblProgress.Text = str1
                '    frmH.lblProgress.Refresh()
                'End If

                var2 = tbl.Rows.Item(Count1).Item("id_tblConfigBodySections")
                strF = "id_tblStudies = " & id_tblStudies & "AND id_tblConfigBodySections = " & var2
                drows = dtbl.Select(strF)
                int1 = drows.Length
                If int1 < 1 Then
                    'add new entry
                    Dim newDRV As DataRow = dtbl.NewRow
                    newDRV("id_tblStudies") = id_tblStudies
                    newDRV("id_tblConfigReportType") = tbl.Rows.Item(Count1).Item("id_tblConfigReportType") 'id_tblConfigReportType
                    var1 = tbl.Rows.Item(Count1).Item("id_tblConfigBodySections")
                    intWordTableNumber = tbl.Rows.Item(Count1).Item("intWordTableNumber")
                    newDRV("id_tblConfigBodySections") = var1 'tbl.Rows.item(Count1).Item("id_tblConfigBodySections")
                    'find 1st charStatement for guwu
                    strF3 = "id_tblConfigBodySections = " & var1
                    Erase rows3
                    rows3 = tbl3.Select(strF3)
                    intRows3 = rows3.Length
                    If intRows3 = 0 Then
                        'must create an entry for tblWordStatements
                        'ID_TBLWORDSTATEMENTS
                        'ID_TBLCONFIGBODYSECTIONS
                        'INTWORDTABLENUMBER
                        'CHARTITLE
                        'CHARWORDSTATEMENT
                        'UPSIZE_TS

                        Dim rowsM() As DataRow
                        Dim strFM As String
                        Dim maxID As Int64
                        maxID = GetMaxID("TBLWORDSTATEMENTS", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
                        '20190219 LEE: Don't need anymore. Used GetMaxID
                        'Call PutMaxID("TBLWORDSTATEMENTS", maxID)

                        'strFM = "CHARTABLE = 'TBLWORDSTATEMENTS'"
                        'rowsM = tblMaxID.Select(strFM)
                        'maxID = rowsM(0).Item("NUMMAXID")
                        'maxID = maxID + 1
                        'rowsM(0).BeginEdit()
                        'rowsM(0).Item("NUMMAXID") = maxID
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


                        Dim newWS As DataRow = tbl3.NewRow
                        newWS.BeginEdit()
                        newWS("ID_TBLWORDSTATEMENTS") = maxID
                        newWS("ID_TBLCONFIGBODYSECTIONS") = var1
                        newWS("INTWORDTABLENUMBER") = intWordTableNumber
                        newWS("CHARTITLE") = "Statement 1"
                        newWS("CHARWORDSTATEMENT") = "Blank" 'this is supposed to be a path. It may be ignored
                        newWS("UPSIZE_TS") = dt
                        newWS.EndEdit()
                        tbl3.Rows.Add(newWS)

                        'Try
                        '    frmH.ta_tblWordStatements.Update(tblWordStatements)
                        'Catch ex As DBConcurrencyException
                        '    'ds2005.TBLWORDSTATEMENTS.Merge('ds2005.TBLWORDSTATEMENTS, True)
                        'End Try

                        var1 = "Statement 1"
                        idWS = maxID

                    Else
                        'var1 = rows3(0).Item("charStatement")
                        'Try

                        'Catch ex As Exception
                        '    var1 = ""

                        'End Try
                        var1 = rows3(0).Item("CHARTITLE")
                        idWS = rows3(0).Item("ID_TBLWORDSTATEMENTS")


                    End If
                    newDRV("charStatement") = var1
                    newDRV("ID_TBLWORDSTATEMENTS") = idWS

                    newDRV("charSectionName") = tbl.Rows.Item(Count1).Item("charSectionName") 'unbound column
                    newDRV("intOrder") = tbl.Rows.Item(Count1).Item("intOrder")
                    newDRV("boolInclude") = 0 'False
                    newDRV("boolUseStatements") = 0 'False
                    newDRV("boolGuWu") = -1 'True
                    newDRV("NUMHEADINGLEVEL") = tbl.Rows.Item(Count1).Item("NUMHEADINGLEVEL")
                    newDRV("NUMCOMPANY") = tbl.Rows.Item(Count1).Item("NUMCOMPANY")
                    newDRV("CHARHEADINGTEXT") = tbl.Rows.Item(Count1).Item("CHARHEADINGTEXT")
                    If tbl.Rows.Item(Count1).Item("NUMHEADINGLEVEL") = 1 Then
                        newDRV("BOOLPAGEBREAK") = -1
                    Else
                        newDRV("BOOLPAGEBREAK") = 0
                    End If
                    dtbl.Rows.Add(newDRV)
                Else
                    drows(0).BeginEdit()
                    var1 = tbl.Rows.Item(Count1).Item("charSectionName") 'unbound column
                    drows(0).Item("charSectionName") = var1
                    'drows(0).Item("intOrder") = tbl.Rows.item(Count1).Item("intOrder")
                    'drows(0).Item("boolInclude") = tbl.Rows.item(Count1).Item("boolInclude")
                    'drows(0).Item("boolUseStatements") = tbl.Rows.item(Count1).Item("boolUseStatements")
                    'drows(0).Item("boolGuWu") = tbl.Rows.item(Count1).Item("boolGuWu")
                    'drows(0).Item("BOOLPAGEBREAK") = tbl.Rows.item(Count1).Item("BOOLPAGEBREAK")
                    'drows(0).Item("CHARHEADINGTEXT") = tbl.Rows.item(Count1).Item("CHARHEADINGTEXT")
                    'drows(0)("NUMCOMPANY") = tbl.Rows.item(Count1).Item("NUMCOMPANY")
                    drows(0).EndEdit()
                End If

            Next

            'keep the following code in case it's needed
            'If ct1 > intRows Then 'add additional rows
            '    For Count1 = intRows To ct1 - 1
            '        'Dim newDRV As DataRowView = dv.AddNew()
            '        Dim newDRV As DataRow = dtbl.NewRow
            '        newDRV("id_tblStudies") = id_tblStudies
            '        newDRV("id_tblConfigReportType") = tbl.Rows.item(Count1).Item("id_tblConfigReportType") 'id_tblConfigReportType
            '        var1 = tbl.Rows.item(Count1).Item("id_tblConfigBodySections")
            '        newDRV("id_tblConfigBodySections") = tbl.Rows.item(Count1).Item("id_tblConfigBodySections")
            '        newDRV("charStatement") = ""
            '        newDRV("charSectionName") = tbl.Rows.item(Count1).Item("charSectionName") 'unbound column
            '        newDRV("intOrder") = tbl.Rows.item(Count1).Item("intOrder")
            '        newDRV("boolInclude") = True
            '        newDRV("boolUseStatements") = False
            '        newDRV("boolGuWu") = True
            '        newDRV("intHeadingLevel") = 1
            '        dtbl.Rows.Add(newDRV)
            '    Next
            'End If
        End If

        If boolFormLoad And Len(NZ(frmH.cbxStudy.Text, "")) = 0 Then
            'first do a datasource of 0 records in order to initialize grid more quickly
            strF = "id_tblStudies = -1"
            Dim dv10 As System.Data.DataView = New DataView(dtbl)
            dv10.RowFilter = strF
            dgv.DataSource = dv10

            int1 = dgv.Rows.Count

            'Dim dt As Date 'for debugging
            'dt = Now
            ''''''''''''''''''console.writeline("Start InitializeReportStatements: " & dt)

            Call InitializeReportStatements()

            'dt = Now
            ''''''''''''''''''console.writeline("End InitializeReportStatements: " & dt)

        End If

        'establish dataview
        If frmH.rbShowIncludedRBody.Checked Then
            strF = "id_tblStudies = " & id_tblStudies & " AND boolInclude = -1 AND NUMCOMPANY <> 20000"
        Else
            strF = "id_tblStudies = " & id_tblStudies & " AND NUMCOMPANY <> 20000"
        End If
        ' strF = "id_tblStudies = " & id_tblStudies & " AND id_tblConfigReportType = " & id_tblConfigReportType
        strS = "intOrder ASC"
        Dim dv As System.Data.DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)

        arrRBSColumns(2, 0) = dv.RowFilter.ToString

        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = True

        'If frmH.rbRBS_Col.Checked Then

        '    var1 = dgv.Columns.Item("CHARHEADINGTEXT").Frozen
        '    dgv.Columns.Item("CHARHEADINGTEXT").Frozen = False
        '    dgv.DataSource = dv
        '    'dgv.Columns.Item("CHARHEADINGTEXT").Frozen = var1
        'Else
        '    var1 = dgv.Columns.Item("charSectionName").Frozen
        '    dgv.Columns.Item("charSectionName").Frozen = False
        '    dgv.DataSource = dv
        '    'dgv.Columns.Item("charSectionName").Frozen = var1
        'End If

        Try
            If frmH.rbRBS_Col.Checked Then

                dgv.DataSource = dv
                var1 = dgv.Columns.Item("CHARHEADINGTEXT").Frozen
                dgv.Columns.Item("CHARHEADINGTEXT").Frozen = False

                'dgv.Columns.Item("CHARHEADINGTEXT").Frozen = var1
            Else
                dgv.DataSource = dv
                var1 = dgv.Columns.Item("charSectionName").Frozen
                dgv.Columns.Item("charSectionName").Frozen = False

                'dgv.Columns.Item("charSectionName").Frozen = var1
            End If
        Catch ex As Exception

        End Try

        frmH.dgvReportStatements.AutoResizeRows()

        dgv.Refresh()

        'select first row
        If frmH.dgvReportStatements.Rows.Count = 0 Then
        Else
            'frmH.dgvReportStatements.CurrentCell = frmH.dgvReportStatements.Rows.Item(0).Cells("charSectionName")
            Try
                frmH.dgvReportStatements.CurrentCell = frmH.dgvReportStatements.Rows.Item(0).Cells("charSectionName")

            Catch ex As Exception

            End Try
        End If
        oldCurrentRowRS = -1
        Call UpdateWord_dgv()

        Call SetRBSCmds()

        'If boolFormLoad Then
        '    frmH.pb1.Value = 0
        '    frmH.pb1.Maximum = omax
        '    frmH.pb1.Value = oval
        '    frmH.pb1.Refresh()
        'End If
        var1 = 1

    End Sub

    Sub InitializeReportStatements()
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView

        dgv = frmH.dgvReportStatements

        dtbl = tblReportstatements

        Dim dt As Date 'for debugging
        'dt =Now
        ''''''''''''''''''console.writeline("Start unfreezing columns: " & dt)

        'int1 = dgv.Rows.Count'for testing
        ''''''''''''''''''console.writeline("Rows: " & int1)


        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).Frozen = False 'for some reason, this has to be done on all columns, or an error gets thrown later
        Next

        'dt =Now
        ''''''''''''''''''console.writeline("End unfreezing columns: " & dt)


        'dt =Now
        ''''''''''''''''''console.writeline("Start displayindex columns: " & dt)

        'all dgv initialization performed in ReportStatementInitialize

        Dim cola As DataColumn
        For Each cola In dtbl.Columns
            str2 = cola.ColumnName
            Select Case str2
                Case "NUMHEADINGLEVEL"
                    cola.Caption = "A*"
                Case "CHARSECTIONNAME"
                    cola.Caption = "Section Name"
                    'cola.ReadOnly = True
                Case "BOOLINCLUDE"
                    cola.Caption = "Include*"
                Case "INTORDER"
                    cola.Caption = "Order"
                Case "BOOLUSESTATEMENTS"
                    cola.Caption = "B*"
                Case "BOOLGUWU"
                    cola.Caption = "C*"
                Case "CHARSTATEMENT"
                    cola.Caption = "Assigned Statement"
                Case "CHARHEADINGTEXT"
                    cola.Caption = "Heading Text"
            End Select
            str1 = cola.Caption
            If InStr(str2, "id_", CompareMethod.Text) > 0 Then
                dgv.Columns.Item(str2).Visible = False
            End If
            dgv.Columns.Item(str2).HeaderText = str1

            'set display index to a large number
            dgv.Columns.Item(str2).DisplayIndex = int1 - 1

        Next

        'dt =Now
        ''''''''''''''''''console.writeline("End displayindex columns: " & dt)

        'dt =Now
        ''''''''''''''''''console.writeline("Start Column Ordering columns: " & dt)

        Call OrderReportStatementCol()

        'dt =Now
        ''''''''''''''''''console.writeline("End Column Ordering columns: " & dt)

        int1 = dgv.Columns.Count
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.ColumnHeader)
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'dgv.SelectionMode = DataGridViewSelectionMode.CellSelect

        'dt =Now
        ''''''''''''''''''console.writeline("Start InitializeReportStatements: " & dt)

        For Count1 = 0 To int1 - 1
            'dgv.Columns.item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            str2 = dgv.Columns.Item(Count1).Name
            Select Case str2
                Case "BOOLPAGEBREAK"
                    dgv.Columns.Item(str2).Visible = False
                Case "boolPB"
                    dgv.Columns.Item(str2).Visible = True
                    dgv.Columns.Item(str2).HeaderText = "B*"
                    dgv.Columns.Item(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgv.Columns.Item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                Case "charSectionName" 'Leave as small caps. column was added later
                    dgv.Columns.Item(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                    dgv.Columns.Item(str2).Width = 250
                    dgv.Columns.Item("charSectionName").ReadOnly = True
                    'dgv.Columns.item("charSectionName").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                    dgv.Columns.Item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                Case "BOOLINCLUDE"
                    dgv.Columns.Item(str2).Visible = False
                    'dgv.Columns.item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                Case "CHARHEADINGTEXT"
                    dgv.Columns.Item(str2).Width = 250
                    dgv.Columns.Item(str2).Visible = True
                    'dgv.Columns.item(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                    dgv.Columns.Item(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                    dgv.Columns.Item("CHARHEADINGTEXT").Visible = True
                    dgv.Columns.Item("CHARHEADINGTEXT").DataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                    dgv.Columns.Item("CHARHEADINGTEXT").DefaultCellStyle.WrapMode = DataGridViewTriState.False
                    dgv.Columns.Item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                Case "NUMHEADINGLEVEL"
                    dgv.Columns.Item(str2).Visible = True
                    dgv.Columns.Item(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgv.Columns.Item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                Case "INTORDER"
                    dgv.Columns.Item("INTORDER").Visible = True
                    dgv.Columns.Item(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgv.Columns.Item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                Case "BOOLUSESTATEMENTS"
                    dgv.Columns.Item("BOOLUSESTATEMENTS").Visible = False
                Case "BOOLGUWU"
                    dgv.Columns.Item(str2).Visible = False
                Case "CHARSTATEMENT"
                    'dgv.Columns.item(str2).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                    dgv.Columns.Item("charStatement").ReadOnly = True
                    dgv.Columns.Item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                Case "UPSIZE_TS"
                    dgv.Columns.Item("UPSIZE_TS").Visible = False
                Case "NUMCOMPANY"
                    dgv.Columns.Item("NUMCOMPANY").Visible = False
                Case "boolUStatements"
                    dgv.Columns.Item(str2).Visible = False
                Case "boolGW"
                    dgv.Columns.Item(str2).Visible = False
                Case "boolI"
                    dgv.Columns.Item(str2).HeaderText = "Include*"
                    dgv.Columns.Item(str2).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv.Columns.Item(str2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            End Select


        Next

        'record dgv column visibility and heading text
        For Count1 = 0 To dgv.Columns.Count - 1
            arrRBSColumns(0, Count1) = dgv.Columns(Count1).Visible
            arrRBSColumns(1, Count1) = dgv.Columns(Count1).HeaderText
        Next

        'dt =Now
        ''''''''''''''''''console.writeline("End InitializeReportStatements: " & dt)

        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        'dgv.Columns.item("charSectionName").Frozen = True
        Dim var1

        var1 = 1

        arrRBSColumns(8, 0) = dgv.Columns("CHARHEADINGTEXT").Width
        arrRBSColumns(9, 0) = dgv.Columns("CHARSTATEMENT").Width
        arrRBSColumns(10, 0) = dgv.AutoSizeColumnsMode

    End Sub

    Sub SetRBSCmds()

        Dim dgv As DataGridView

        'set cmdOrderReportBody and cmdRSBAll position
        Dim wd1, wd2

        Try
            'Dim dgv As DataGridView
            dgv = frmH.dgvReportStatements
            wd1 = 0
            wd2 = 0
            wd2 = dgv.RowHeadersWidth
            wd1 = wd1 + wd2
            wd2 = dgv.Columns.Item("NUMHEADINGLEVEL").Width
            wd1 = wd1 + wd2
            wd2 = dgv.Columns.Item("boolPB").Width
            wd1 = wd1 + wd2
            wd2 = dgv.Columns.Item("charSectionName").Width
            wd1 = wd1 + wd2
            frmH.cmdRBSAll.Left = dgv.Left + wd1
            wd2 = dgv.Columns.Item("boolInclude").Width
            wd1 = wd1 + wd2
            frmH.cmdOrderReportBodySection.Left = dgv.Left + wd1
            'dgv.Columns.item("intOrder").Width = frmh.cmdOrderReportBodySection.Width
        Catch ex As Exception

        End Try

    End Sub

    Sub ReportStatementsFillCharSection()
        Dim tbl As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim ct1 As Short
        Dim Count1 As Short
        Dim var1, var2, var3
        Dim strF As String
        Dim int1 As Short
        Dim drows() As DataRow
        Dim bool As Boolean
        Dim Count2 As Short
        Dim str1 As String
        Dim str2 As String

        'this routine will also fill bools


        tbl = tblConfigBodySections
        ct1 = tbl.Rows.Count
        dtbl = tblReportstatements

        'do next only on getstudyinfo or reset Action
        If boolRSCFill Then
            For Count1 = 0 To ct1 - 1

                var2 = tbl.Rows.Item(Count1).Item("id_tblConfigBodySections")
                strF = "id_tblStudies = " & id_tblStudies & "AND id_tblConfigBodySections = " & var2
                Erase drows
                drows = dtbl.Select(strF)
                int1 = drows.Length

                If int1 = 0 Then
                Else
                    drows(0).BeginEdit()
                    var1 = tbl.Rows.Item(Count1).Item("charSectionName") 'unbound column
                    drows(0).Item("charSectionName") = var1
                    'For Count2 = 1 To 3
                    '    Select Case Count2
                    '        Case 1
                    '            str1 = "boolInclude"
                    '            str2 = "boolI"
                    '        Case 2
                    '            str1 = "boolUseStatements"
                    '            str2 = "boolUStatements"
                    '        Case 3
                    '            str1 = "boolGuWu"
                    '            str2 = "boolGW"
                    '    End Select
                    '    int1 = drows(0).Item(str1)
                    '    If int1 = -1 Then
                    '        bool = True
                    '    Else
                    '        bool = False
                    '    End If
                    '    drows(0).Item(str2) = bool
                    'Next

                    str1 = "boolInclude"
                    str2 = "boolI"
                    int1 = drows(0).Item(str1)
                    If int1 = -1 Then
                        bool = True
                    Else
                        bool = False
                    End If
                    drows(0).Item(str2) = bool

                    str1 = "BOOLPAGEBREAK"
                    str2 = "boolPB"
                    int1 = drows(0).Item(str1)
                    If int1 = -1 Then
                        bool = True
                    Else
                        bool = False
                    End If
                    drows(0).Item(str2) = bool

                    drows(0).EndEdit()
                End If

            Next
        End If

    End Sub

    Sub OrderReportStatementCol()

        Try
            frmH.dgvReportStatements.Columns.Item("charSectionName").Frozen = False
            frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").Frozen = False
            frmH.dgvReportStatements.AllowUserToOrderColumns = True
            'now order the columns properly
            frmH.dgvReportStatements.Columns.Item("NUMHEADINGLEVEL").DisplayIndex = 0
            frmH.dgvReportStatements.Columns.Item("boolPB").DisplayIndex = 1
            frmH.dgvReportStatements.Columns.Item("charSectionName").ReadOnly = True
            frmH.dgvReportStatements.Columns.Item("boolI").DisplayIndex = 3
            frmH.dgvReportStatements.Columns.Item("intOrder").DisplayIndex = 4
            'frmh.dgvReportStatements.Columns.item("boolUStatements").DisplayIndex = 4
            'frmh.dgvReportStatements.Columns.item("boolGW").DisplayIndex = 5
            frmH.dgvReportStatements.Columns.Item("charStatement").DisplayIndex = 5
            If frmH.rbRBS_Col.Checked Then
                frmH.dgvReportStatements.Columns.Item("charSectionName").DisplayIndex = 6
                frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").DisplayIndex = 2
                'frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").Frozen = True
            Else
                frmH.dgvReportStatements.Columns.Item("charSectionName").DisplayIndex = 2
                frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").DisplayIndex = 6
                'frmH.dgvReportStatements.Columns.Item("charSectionName").Frozen = True
            End If
            frmH.dgvReportStatements.AllowUserToOrderColumns = False
            'frmh.dgvReportStatements.Columns.item("charSectionName").Frozen = True

        Catch ex As Exception

        End Try

        Call frmH.ViewSections(False)

    End Sub

    Sub DoReportStatementsCancel()

        boolRSCFill = True

        Call SetEntireReportRButton()
        Call ReportStatementsFill()
        Call ReportStatementsFillCharSection()
        Call RBFilter()

        boolRSCFill = False

    End Sub

    Sub SetEntireReportRButton()
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim intL As Short
        Dim var1

        strF = "ID_TBLSTUDIES = " & id_tblStudies
        dtbl = tblData
        rows = dtbl.Select(strF)
        intL = rows.Length

        If intL = 0 Then
            frmH.rbEntireReport.Checked = True
        Else
            var1 = rows(0).Item("BOOLENTIREREPORT")
            If var1 = -1 Then
                frmH.rbEntireReport.Checked = True
            Else
                frmH.rbSections.Checked = True
            End If
        End If

    End Sub

    Sub CPTab_Initialize()
        Dim tbl As System.Data.DataTable
        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim drow As DataRow
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim Count1 As Short
        Dim boolg As Boolean
        Dim strT As String
        Dim strF As String
        Dim var1, var2
        Dim wi As Short
        Dim twi As Short
        Dim col2 As DataColumn
        Dim col As DataColumn
        Dim strS As String
        Dim boolV As Boolean
        Dim int1 As Short

        tbl = tblCP
        dgv = frmH.dgvContributingPersonnel
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dtbl = tblContributingPersonnel

        'add one unbound columns
        Dim col1 As New DataColumn
        str1 = "boolIncludeSOTP"
        str2 = "A_*"
        col1.ColumnName = str1
        col1.Caption = str2
        col1.DataType = System.Type.GetType("System.Boolean")
        col1.DefaultValue = False
        col1.AllowDBNull = False
        dtbl.Columns.Add(col1)

        'For Count1 = 1 To 3
        '    Dim col1 As New DataColumn
        '    Select Case Count1
        '        Case 1
        '            str1 = "boolIncludeSOP"
        '            str2 = "A_*"
        '        Case 2
        '            str1 = "boolIncludeSOTP"
        '            str2 = "B_*"
        '        Case 3
        '            str1 = "boolIncludeSOCS"
        '            str2 = "C_*"
        '    End Select
        '    col1.ColumnName = str1
        '    col1.Caption = str2
        '    col1.DataType = System.Type.GetType("System.Boolean")
        '    col1.DefaultValue = False
        '    col1.AllowDBNull = False
        '    dtbl.Columns.Add(col1)
        'Next

        strF = "id_tblStudies = " & id_tblStudies
        strS = "intOrder ASC"
        Dim dv As System.Data.DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowNew = False
        dgv.DataSource = dv

        'make all columns invisible
        For Count1 = 0 To dtbl.Columns.Count - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable

        Next

        Dim strCol As String

        twi = CInt(dgv.Width)
        wi = 0
        int1 = dgv.Columns.Count
        Dim tl As Short
        Dim ctl As Control
        Dim ctl1 As Control
        tl = dgv.RowHeadersWidth + dgv.Left
        For Count1 = 1 To 8
            boolg = False
            boolV = False
            Select Case Count1
                Case 1
                    str3 = "charCPPrefix"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Prefix"
                    wi = twi * 0.05
                Case 2
                    str3 = "charCPName"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Name **"
                    wi = twi * 0.2
                Case 3
                    str3 = "charCPSuffix"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Suffix"
                    wi = twi * 0.05
                Case 4
                    str3 = "charCPDegree"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Degree"
                    wi = twi * 0.06
                Case 5
                    str3 = "charCPTitle"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Title"
                    wi = twi * 0.1675
                Case 6
                    str3 = "charCPRole"
                    boolg = True
                    boolV = True
                    str1 = "text"
                    str2 = "Role"
                    wi = twi * 0.1675
                    'Case 7
                    '    str3 = "boolIncludeSOP"
                    '    boolg = False
                    '    boolV = True
                    '    str1 = "bool"
                    '    str2 = "A *"
                    '    wi = twi * 0.04
                Case 7
                    str3 = "boolIncludeSOTP"
                    boolg = False
                    boolV = True
                    str1 = "bool"
                    str2 = "A *"
                    wi = twi * 0.04
                    'Case 9
                    '    str3 = "boolIncludeSOCS"
                    '    boolg = False
                    '    boolV = True
                    '    str1 = "bool"
                    '    str2 = "C *"
                    '    wi = twi * 0.04
                Case 8
                    str3 = "intOrder"
                    boolg = False
                    boolV = True
                    str1 = "textc"
                    str2 = "B *"
                    wi = twi * 0.04

            End Select

            If boolV Then
                dgv.Columns.Item(str3).Visible = True
                dgv.Columns.Item(str3).HeaderText = str2
                dgv.Columns.Item(str3).Width = wi
                dgv.Columns.Item(str3).MinimumWidth = wi
            End If
            If StrComp(str3, "intOrder", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(str3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

        Next

        'set displayorder
        Call CPDisplayOrder()

        Call FillDropdownBoxes()

        Call CP_FillTable()

    End Sub

    Sub CPDisplayOrder()
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim str3 As String

        dgv = frmH.dgvContributingPersonnel
        int1 = dgv.Columns.Count
        'first set everything to a high display order
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).DisplayIndex = int1 - 1
        Next

        'set display order
        For Count1 = 0 To int1 - 1
            str3 = ""
            Select Case Count1
                Case 1
                    str3 = "charCPPrefix"
                    int2 = 0
                Case 2
                    str3 = "charCPName"
                    int2 = 1
                Case 3
                    str3 = "charCPSuffix"
                    int2 = 2
                Case 4
                    str3 = "charCPDegree"
                    int2 = 3
                Case 5
                    str3 = "charCPTitle"
                    int2 = 4
                Case 6
                    str3 = "charCPRole"
                    int2 = 5
                    'Case 7
                    '    str3 = "boolIncludeSOP"
                    '    int2 = 6
                Case 7
                    str3 = "boolIncludeSOTP"
                    int2 = 6
                    'Case 9
                    '    str3 = "boolIncludeSOCS"
                    '    int2 = 8
                Case 8
                    str3 = "intOrder"
                    int2 = 7

            End Select
            If Len(str3) = 0 Then
            Else
                dgv.Columns.Item(str3).DisplayIndex = int2
            End If
        Next

    End Sub

    Sub FillHomeDropdownBoxes()

        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim var1
        Dim rows() As DataRow

        boolStopCBX = True

        'retrieve report types
        Dim int1 As Short
        tbl = tblConfigReportType
        strF = "ID_TBLCONFIGREPORTTYPE < 1000"
        strS = "ID_TBLCONFIGREPORTTYPE ASC"
        rows = tbl.Select(strF, strS)
        int1 = rows.Length
        'var1 = frmh.cbxCPDegree.SelectedIndex
        cbxxReportTypes.Items.Clear()
        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("CHARREPORTTYPE"), "Sample Analysis")
            cbxxReportTypes.Items.Add(str1)
        Next
        cbxxReportTypes.AutoComplete = True
        cbxxReportTypes.MaxDropDownItems = 20
        cbxxReportTypes.DisplayStyleForCurrentCellOnly = True
        cbxxReportTypes.DropDownWidth = cbxxReportTypes.DropDownWidth * 1.5
        cbxxReportTypes.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

    End Sub

    Sub FillDropdownBoxes()

        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim var1

        boolStopCBX = True

        'retrieve contributing personnel Degree
        Dim rows() As DataRow
        Dim strS As String
        Dim int1 As Short
        tbl = tblDropdownBoxContent
        Erase rows
        strF = "id_tblDropdownBoxName = 4"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        cbxxCPDegree.Items.Clear()
        cbxxCPDegree.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPDegree.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("charValue"), "")
            cbxxCPDegree.Items.Add(str1)
        Next
        'cbxxCPDegree.Value = ""
        cbxxCPDegree.AutoComplete = True
        cbxxCPDegree.MaxDropDownItems = 20
        cbxxCPDegree.Sorted = True
        cbxxCPDegree.DisplayStyleForCurrentCellOnly = True
        cbxxCPDegree.DropDownWidth = cbxxCPDegree.DropDownWidth * 2
        cbxxCPDegree.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel title
        Erase rows
        strF = "id_tblDropdownBoxName = 8"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        cbxxCPTitle.Items.Clear()
        cbxxCPTitle.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("charValue"), "")
            cbxxCPTitle.Items.Add(str1)
        Next
        'cbxxCPTitle.Value = ""
        cbxxCPTitle.AutoComplete = True
        cbxxCPTitle.MaxDropDownItems = 20
        cbxxCPTitle.Sorted = True
        cbxxCPTitle.DisplayStyleForCurrentCellOnly = True
        cbxxCPTitle.DropDownWidth = cbxxCPTitle.DropDownWidth * 2
        cbxxCPTitle.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel prefix
        Erase rows
        strF = "id_tblDropdownBoxName = 5"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPPrefix.Items.Add("[None]")
        cbxxCPPrefix.Items.Clear()
        cbxxCPPrefix.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charValue")
            cbxxCPPrefix.Items.Add(str1)
        Next
        'cbxxCPPrefix.Value = ""
        cbxxCPPrefix.AutoComplete = True
        cbxxCPPrefix.MaxDropDownItems = 20
        cbxxCPPrefix.DisplayStyleForCurrentCellOnly = True
        cbxxCPPrefix.DropDownWidth = cbxxCPPrefix.DropDownWidth * 2
        cbxxCPPrefix.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel suffix
        Erase rows
        strF = "id_tblDropdownBoxName = 6"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPSuffix.Items.Add("[None]")
        cbxxCPSuffix.Items.Clear()
        cbxxCPSuffix.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charValue")
            cbxxCPSuffix.Items.Add(str1)
        Next
        'cbxxCPSuffix.Value = ""
        cbxxCPSuffix.AutoComplete = True
        cbxxCPSuffix.MaxDropDownItems = 20
        cbxxCPSuffix.DisplayStyleForCurrentCellOnly = True
        cbxxCPSuffix.DropDownWidth = cbxxCPSuffix.DropDownWidth * 2
        cbxxCPSuffix.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve contributing personnel role
        Erase rows
        strF = "id_tblDropdownBoxName = 7"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPRole.Items.Add("[None]")
        cbxxCPRole.Items.Clear()
        cbxxCPRole.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = NZ(rows(Count1).Item("charValue"), "")
            'If Len(str1) = 0 Then
            'Else
            '    cbxxCPRole.Items.Add(str1)
            'End If
            cbxxCPRole.Items.Add(str1)
        Next
        'cbxxCPRole.Value = ""
        cbxxCPRole.AutoComplete = True
        cbxxCPRole.MaxDropDownItems = 20
        cbxxCPRole.Sorted = True
        cbxxCPRole.DisplayStyleForCurrentCellOnly = True
        cbxxCPRole.DropDownWidth = cbxxCPRole.DropDownWidth * 2
        cbxxCPRole.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        'retrieve Assay Proc Descr
        Erase rows
        strF = "id_tblDropdownBoxName = 12"
        strS = "intOrder ASC"
        rows = tbl.Select(strF, strS)
        'frmh.cbxCPRole.Items.Add("[None]")
        cbxxAssayDescr.Items.Clear()
        'cbxxAssayDescr.Items.Add("")
        int1 = rows.Length
        'var1 = frmh.cbxCPTitle.SelectedIndex
        For Count1 = 0 To int1 - 1
            str1 = rows(Count1).Item("charValue")
            cbxxAssayDescr.Items.Add(str1)
        Next
        'cbxxCPRole.Value = ""
        cbxxAssayDescr.AutoComplete = True
        cbxxAssayDescr.MaxDropDownItems = 20
        cbxxAssayDescr.Sorted = True
        cbxxAssayDescr.DisplayStyleForCurrentCellOnly = True
        cbxxAssayDescr.DropDownWidth = cbxxCPRole.DropDownWidth * 2
        cbxxAssayDescr.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        cbxxAnalMethType.Items.Clear()
        For Count1 = 0 To frmH.cbxAssayTechniqueAcronym.Items.Count - 1
            str1 = NZ(frmH.cbxAssayTechniqueAcronym.Items(Count1), "NA")
            cbxxAnalMethType.Items.Add(str1)
        Next
        cbxxAnalMethType.AutoComplete = True
        cbxxAnalMethType.MaxDropDownItems = 20
        cbxxAnalMethType.Sorted = True
        cbxxAnalMethType.DisplayStyleForCurrentCellOnly = True
        'cbxxAnalMethType.DropDownWidth = cbxxCPRole.DropDownWidth * 1.5
        cbxxAnalMethType.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton


        'enter yes/no
        'frmh.cbxCPCoverPageSigBlock.Items.Add("Yes")
        'frmh.cbxCPCoverPageSigBlock.Items.Add("No")
        'frmh.cbxCPCoverPageSigBlock.SelectedIndex = 1
        'frmh.cbxCPTableSigBlock.Items.Add("Yes")
        'frmh.cbxCPTableSigBlock.Items.Add("No")
        'frmh.cbxCPTableSigBlock.SelectedIndex = 1

        'enter Personnel
        tbl = tblPersonnel
        Dim tdv As System.Data.DataView = New DataView(tbl)
        'REDO THIS LINE!!!!
        'tdv = tbl.DefaultView
        tdv.Sort = "charLastName ASC"
        int1 = tdv.Count
        cbxxCPName.Items.Clear()
        cbxxCPName.Items.Add("")
        For Count1 = 0 To int1 - 1
            str1 = tdv(Count1).Item("charFIRSTNAME")
            str2 = NZ(tdv(Count1).Item("charMIDDLEname"), "")
            str3 = tdv(Count1).Item("charLASTNAME")
            If StrComp(str3, "aaAdmin", CompareMethod.Text) = 0 Then
            Else
                If Len(str2) = 0 Then 'no middle initial provided
                    str4 = str1 & " " & str3
                Else
                    If Len(str2) = 1 Then 'needs a period
                        str2 = str2 & "."
                    Else
                    End If
                    str4 = str1 & " " & str2 & " " & str3
                End If
                cbxxCPName.Items.Add(str4)
            End If
        Next
        int1 = rows.Length
        'cbxxCPName.Value = ""
        cbxxCPName.AutoComplete = True
        cbxxCPName.MaxDropDownItems = 20
        cbxxCPName.DisplayStyleForCurrentCellOnly = True
        cbxxCPName.DropDownWidth = cbxxCPName.DropDownWidth * 1.5
        cbxxCPName.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        boolStopCBX = False

    End Sub

    Sub Update_cbxAnalytes()

        Dim strF As String
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim var1
        Dim dv As System.Data.DataView
        Dim dgv As DataGridView

        'retrieve analytes from frmh.dgvCompanyAnalRef
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short

        dgv = frmH.dgvCompanyAnalRef
        dv = dgv.DataSource
        int1 = dgv.Columns.Count
        int2 = FindRowDVByCol("Is Replicate?", dv, "Item")
        int3 = FindRowDVByCol("Analyte Name", dv, "Item")

        boolStopCBX = True

        'in case there is a hook running, this information must come from dgv dataview
        cbxAnalytes.Items.Clear()
        cbxAnalytes.Items.Add("")
        For Count1 = 0 To int1 - 1
            var1 = dgv.Columns.Item(Count1).Name
            If StrComp(var1, "Item", CompareMethod.Text) = 0 Then 'ignore
            ElseIf StrComp(var1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'ignore
            ElseIf StrComp(var1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'ignore
            Else
                'record analyte name
                str1 = NZ(dv(int3).Item(var1), "") 'Analyte Name
                str2 = NZ(dv(int2).Item(var1), "") 'Is Replicate?
                If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                Else
                    cbxAnalytes.Items.Add(str1)
                End If
            End If
        Next
        cbxAnalytes.AutoComplete = True
        cbxAnalytes.MaxDropDownItems = 20
        cbxAnalytes.DisplayStyleForCurrentCellOnly = True
        cbxAnalytes.DropDownWidth = cbxAnalytes.DropDownWidth * 1.25
        cbxAnalytes.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton

        boolStopCBX = False

    End Sub

    Sub CP_Reset()
        Dim dtbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim strF As String
        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim strS As String

        tbl = tblCP
        dgv = frmH.dgvContributingPersonnel
        dtbl = tblContributingPersonnel
        dtbl.RejectChanges()

        Call UpdateCPBool()
        'strF = "id_tblStudies = " & id_tblStudies
        'dv = dtbl.DefaultView
        'dv.RowFilter = strF
        'strS = "intOrder ASC"
        'dv.Sort = strS
        'dv.RowStateFilter = DataViewRowState.CurrentRows

        'dv.AllowNew = False
        'dv.AllowEdit = True
        'dv.AllowDelete = False
        'dg.DataSource = dv
        'dg.Refresh()
    End Sub

    Sub CP_FillTable()
        Dim dtbl As System.Data.DataTable
        'Dim dv as system.data.dataview
        Dim strF As String
        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim strS As String

        tbl = tblCP
        dgv = frmH.dgvContributingPersonnel
        dtbl = tblContributingPersonnel
        dtbl.AcceptChanges()
        strF = "id_tblStudies = " & id_tblStudies
        strS = "intOrder ASC"
        Dim dv As System.Data.DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowNew = False
        dv.AllowEdit = True
        dv.AllowDelete = False
        dgv.DataSource = dv

        Call UpdateCPBool()

        dgv.Refresh()

    End Sub

    Sub UpdateCPBool()
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int2 As Short
        Dim bool As Boolean
        Dim str1 As String
        Dim str2 As String

        dv = frmH.dgvContributingPersonnel.DataSource
        int1 = dv.Count
        For Count1 = 0 To int1 - 1
            dv(Count1).BeginEdit()
            str1 = "boolIncludeSOTP"
            str2 = "boolIncludeSigOnTablePage"
            int2 = NZ(dv(Count1).Item(str2), 0)
            If int2 = -1 Then
                bool = True
            Else
                bool = False
            End If
            dv(Count1).Item(str1) = bool
            dv(Count1).EndEdit()

            'For Count2 = 1 To 1
            '    Select Case Count2
            '        Case 1
            '            str1 = "boolIncludeSOP"
            '            str2 = "boolIncludeSigOnCoverPage"
            '        Case 2
            '            str1 = "boolIncludeSOTP"
            '            str2 = "boolIncludeSigOnTablePage"
            '        Case 3
            '            str1 = "boolIncludeSOCS"
            '            str2 = "boolIncludeSigOnCompStatement"
            '    End Select
            '    int2 = dv(Count1).Item(str2)
            '    If int2 = -1 Then
            '        bool = True
            '    Else
            '        bool = False
            '    End If
            '    dv(Count1).Item(str1) = bool
            'Next

            'dv(Count1).EndEdit()
        Next

    End Sub

    Sub doCPCancel()

        Call CP_Reset()

    End Sub

    'Sub ReportHistoryInitialize()
    '    Dim dv as system.data.dataview
    '    Dim var1, var2
    '    Dim str1 As String
    '    Dim ts1 As New DataGridTableStyle
    '    Dim i As Short
    '    Dim gs As DataGridColumnStyle

    '    Dim dgt As DataGridTableStyle
    '    'ts1.MappingName = frmh.dgvReports.DataMember.ToString 'var1
    '    ts1.MappingName = tblReportHistory.TableName
    '    ts1.AllowSorting = False
    '    Dim dgc1 As New DataGridTextBoxColumn
    '    dgc1.MappingName = "dtReportGenerated"
    '    dgc1.HeaderText = "Date"
    '    'dgc1.Format = "mm/dd/yyyy hh:mm:ss AM/PM"
    '    'dgc1.Alignment = HorizontalAlignment.Center
    '    dgc1.NullText = ""
    '    ts1.GridColumnStyles.Add(dgc1)
    '    Dim dgc2 As New DataGridTextBoxColumn
    '    dgc2.MappingName = "charReportGeneratedStatus"
    '    dgc2.HeaderText = "Type"
    '    'dgc2.Alignment = HorizontalAlignment.Center
    '    dgc2.NullText = ""
    '    ts1.GridColumnStyles.Add(dgc2)
    '    frmH.dgReportHistory.TableStyles.Add(ts1)

    '    dv = tblReportHistory.DefaultView
    '    dv.RowFilter = "id_tblReportHistory = 0"
    '    frmH.dgReportHistory.DataSource = dv
    '    frmH.dgReportHistory.TableStyles(0).RowHeaderWidth = 10
    '    frmH.dgReportHistory.CaptionText = "In  Desc.  Order"
    '    frmH.dgReportHistory.Refresh()

    '    'Call AutoSizeGrid(250, tblReportHistory, frmh.dgReportHistory, dv.Count, 2, 0, False)


    'End Sub

    'Sub SetReportHistory()
    '    Dim var1
    '    Dim int1 As Short
    '    Dim strF As String
    '    Dim dv as system.data.dataview
    '    Dim dv1 as system.data.dataview

    '    ''int1 = frmh.dgvReports.CurrentRow.Index
    '    'If int1 = -1 Then 'no reports configured or selected
    '    '    Exit Sub
    '    'End If
    '    If frmH.dgvReports.CurrentRow Is Nothing Then
    '        Exit Sub
    '    End If

    '    int1 = frmH.dgvReports.CurrentRow.Index
    '    If int1 = -1 Then 'no reports configured or selected
    '        Exit Sub
    '    End If
    '    dv = frmH.dgvReports.DataSource
    '    var1 = dv.Item(int1).Item("id_tblReports")
    '    strF = "id_tblStudies = " & id_tblStudies & " and id_tblReports = " & var1

    '    dv1 = tblReportHistory.DefaultView
    '    dv1.RowFilter = strF
    '    dv1.Sort = "id_tblReportHistory DESC"
    '    frmH.dgReportHistory.DataSource = dv1

    '    'Call AutoSizeGrid(250, tblReportHistory, frmh.dgReportHistory, dv.Count, 2, 0, False)

    'End Sub


    Sub ClearForm()

        Dim cl As Control
        Dim var1, var2
        Dim ct1 As Short
        Dim ct2 As Short
        Dim ct3 As Short
        Dim tp As TabPage

        For Each cl In frmH.Controls
            var1 = Microsoft.VisualBasic.Left(cl.Name, 4)
            If StrComp(var1, "char", CompareMethod.Text) = 0 Then
                cl.Text = ""
            End If
            var1 = Microsoft.VisualBasic.Left(cl.Name, 3)
            If StrComp(var1, "cbx", CompareMethod.Text) = 0 Then
                If StrComp(cl.Name, "cbxStudy", CompareMethod.Text) = 0 Then
                ElseIf StrComp(cl.Name, "cbxHooks", CompareMethod.Text) = 0 Then
                ElseIf StrComp(cl.Name, "cbxExampleReport", CompareMethod.Text) = 0 Then
                Else
                    cl.Text = "[None]"
                End If
            End If
        Next

        For Each tp In frmH.tab1.TabPages
            For Each cl In tp.Controls
                If StrComp(cl.Name, "txtcbxMDBSelIndex", CompareMethod.Text) = 0 Then
                Else
                    var1 = Microsoft.VisualBasic.Left(cl.Name, 4)
                    If StrComp(var1, "char", CompareMethod.Text) = 0 Then
                        cl.Text = ""
                    End If
                    var1 = Microsoft.VisualBasic.Left(cl.Name, 3)
                    If StrComp(var1, "cbx", CompareMethod.Text) = 0 Then
                        boolLoad = True
                        cl.Text = "[None]"
                        boolLoad = False
                    End If
                    var1 = Microsoft.VisualBasic.Left(cl.Name, 3)
                    If StrComp(var1, "txt", CompareMethod.Text) = 0 And StrComp(cl.Name, "txt1RBS", CompareMethod.Text) <> 0 Then
                        cl.Text = ""
                    End If
                End If
            Next
        Next

    End Sub

    Sub CreateTables()

        Dim wid1 As Single
        Dim wid2 As Single
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strTable As String
        Dim strDTable As String
        Dim dtbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim dg As DataGridView
        Dim dgv As DataGridView
        Dim boolRO As Boolean
        'Dim 'rsdt As New ADODB.Recordset
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim drow As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2, var3
        Dim dvDT As System.Data.DataView
        Dim strS As String

        dvDT = tblDataTableRowTitles.DefaultView

        Try
            For Count1 = 1 To 5
                Select Case Count1
                    Case 1
                        strDTable = "tblCompanyAnalRefTable"
                        tblCompanyAnalRefTable = New System.Data.DataTable
                        dtbl = tblCompanyAnalRefTable
                        'dg = frmh.dgCompanyAnalRef
                        boolRO = False
                    Case 2
                        strDTable = "tblWatsonAnalRefTable"
                        tblWatsonAnalRefTable = New System.Data.DataTable
                        dtbl = tblWatsonAnalRefTable
                        'dg = frmh.dgWatsonAnalRef
                        boolRO = True
                    Case 3
                        strDTable = "tblWatsonData"
                        tblWatsonData = New System.Data.DataTable
                        dtbl = tblWatsonData
                        dg = frmH.dgvDataWatson
                        boolRO = True
                    Case 4
                        strDTable = "tblCompanyData"
                        tblCompanyData = New System.Data.DataTable
                        dtbl = tblCompanyData
                        'dg = frmh.dgDataCompany
                        boolRO = True

                    Case 5
                        Try
                            strDTable = "tblMethodValidationData"
                        Catch ex As Exception
                            MsgBox("strDTable = " & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dtbl = tblMethodValData
                        Catch ex As Exception
                            MsgBox("dtbl = tblMethodValData " & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv = frmH.dgvMethodValData
                        Catch ex As Exception
                            MsgBox(" dgv = frmH.dgvMethodValData " & ChrW(10) & ex.Message)
                        End Try
                        boolRO = True

                End Select

                Try
                    str1 = "charDataTableName = '" & strDTable & "' AND BOOLINCLUDE <> 0"
                    'rsdt.Filter = str1
                    'rsdt.Sort = "intOrder"
                    dvDT.RowFilter = ""
                    dvDT.RowFilter = str1
                    dvDT.Sort = "intOrder ASC"
                Catch ex As Exception
                    MsgBox("str1 = 'charDataTableName = '' & strDTable & '" & ChrW(10) & ex.Message)
                End Try


                'Dim ts1 As New DataGridTableStyle
                'create a new datagridtablestyle
                'configure datagridtablestyle 

                Dim tcol1 As New DataColumn
                Dim tcol2 As New DataColumn

                If StrComp(strDTable, "tblCompanyAnalRefTable", CompareMethod.Text) = 0 Then
                    Dim tcol3 As New DataColumn
                    Dim tcol4 As New DataColumn

                    'add an ID column
                    tcol3.DataType = System.Type.GetType("System.Int64")
                    tcol3.ColumnName = "ID_TBLDATATABLEROWTITLES"
                    tcol3.Caption = "ID"
                    tcol3.ReadOnly = False
                    dtbl.Columns.Add(tcol3)

                    'add a boolinclude column
                    tcol4.DataType = System.Type.GetType("System.Boolean")
                    tcol4.ColumnName = "BOOLINCLUDE"
                    tcol4.Caption = "A*"
                    tcol4.ReadOnly = False
                    tcol4.DefaultValue = False
                    dtbl.Columns.Add(tcol4)
                End If

                Try
                    tcol1.DataType = System.Type.GetType("System.String")
                    tcol1.ColumnName = "Item"
                    tcol1.Caption = "Item"
                    tcol1.ReadOnly = True
                    dtbl.Columns.Add(tcol1)
                Catch ex As Exception
                    MsgBox("tcol1.DataType = " & ChrW(10) & ex.Message)
                End Try


                'If Count1 = 1 Or Count1 = 2 Then
                'Else
                Try
                    tcol2.DataType = System.Type.GetType("System.String")
                    tcol2.ColumnName = "Value"
                    tcol2.Caption = "Value"
                    tcol2.ReadOnly = False
                    dtbl.Columns.Add(tcol2)
                Catch ex As Exception
                    MsgBox("tcol2.DataType = " & ChrW(10) & ex.Message)
                End Try

                'End I

                Try
                    For Count2 = 0 To dvDT.Count - 1
                        'str1 = rsdt.Fields("charRowName").Value
                        str1 = dvDT(Count2).Item("charRowName")
                        var1 = dvDT(Count2).Item("ID_TBLDATATABLEROWTITLES")
                        '''''''''''''''''''''''''''console.writeline(strDTable & ": " & str1)
                        drow = dtbl.NewRow()
                        '.Rows.item(int1).Item(1) = wStudyID
                        drow.Item("Item") = str1
                        If Count1 = 1 Then
                            drow.Item("ID_TBLDATATABLEROWTITLES") = var1
                            drow.Item("BOOLINCLUDE") = False
                        ElseIf Count1 = 2 Then 'ignore
                        Else
                            drow.Item(1) = ""
                        End If
                        dtbl.Rows.Add(drow)
                    Next
                Catch ex As Exception
                    MsgBox("For Count2 = 0 To dvDT.Count - 1" & ChrW(10) & ex.Message)
                End Try


                'var1 = dg.TableStyles(0).GridColumnStyles
                'refresh datagrid

                'dv = dtbl.DefaultView
                Try
                    dv = New DataView(dtbl)
                    dv.AllowDelete = False
                    dv.AllowNew = False
                    dv.AllowEdit = True
                Catch ex As Exception
                    MsgBox("dv = New DataView(dtbl)" & ChrW(10) & ex.Message)
                End Try

                If Count1 = 1 Then
                    frmH.dgvCompanyAnalRef.DataSource = dv
                    frmH.dgvCompanyAnalRef.Columns.Item("ID_TBLDATATABLEROWTITLES").Visible = False
                    frmH.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").Visible = True
                    frmH.dgvCompanyAnalRef.Columns.Item("ID_TBLDATATABLEROWTITLES").HeaderText = "ID"
                    frmH.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
                    frmH.dgvCompanyAnalRef.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                    int1 = frmH.dgvCompanyAnalRef.Columns.Count
                    For Count2 = 0 To int1 - 1
                        frmH.dgvCompanyAnalRef.AutoResizeColumn(Count2, DataGridViewAutoSizeColumnMode.AllCells)
                        frmH.dgvCompanyAnalRef.Columns.Item(Count2).SortMode = DataGridViewColumnSortMode.NotSortable
                    Next
                    frmH.dgvCompanyAnalRef.AllowUserToResizeColumns = True
                    frmH.dgvCompanyAnalRef.AllowUserToResizeRows = True
                    frmH.dgvCompanyAnalRef.RowHeadersWidth = 25
                    frmH.dgvCompanyAnalRef.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                    ''make last row readonly
                    int1 = frmH.dgvCompanyAnalRef.Rows.Count 'debugging
                    int2 = int1 'debugging
                    'now make some rows invisible
                    Call HideAnalRefRows()
                    'frmH.dgvCompanyAnalRef.Rows.item(int1 - 1).ReadOnly = True
                ElseIf Count1 = 2 Then
                    frmH.dgvWatsonAnalRef.DataSource = dv
                    frmH.dgvWatsonAnalRef.Columns.Item(0).Frozen = True
                    frmH.dgvWatsonAnalRef.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                    int1 = frmH.dgvWatsonAnalRef.Columns.Count
                    For Count2 = 0 To int1 - 1
                        frmH.dgvWatsonAnalRef.AutoResizeColumn(Count2, DataGridViewAutoSizeColumnMode.AllCells)
                        frmH.dgvWatsonAnalRef.Columns.Item(Count2).SortMode = DataGridViewColumnSortMode.NotSortable
                    Next
                    frmH.dgvWatsonAnalRef.AllowUserToResizeColumns = True
                    frmH.dgvWatsonAnalRef.AllowUserToResizeRows = True
                    frmH.dgvWatsonAnalRef.RowHeadersWidth = 25
                    frmH.dgvWatsonAnalRef.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                    'make last row readonly
                    'int1 = frmH.dgvWatsonAnalRef.Rows.Count
                    'frmH.dgvWatsonAnalRef.Rows.item(int1 - 1).ReadOnly = True

                ElseIf Count1 = 3 Then 'tblWatsonData

                    Try
                        dg.DataSource = dv
                        'dg.DataSource = dtbl
                        dg.Refresh()

                        dg.AllowUserToResizeColumns = True
                        dg.AllowUserToResizeRows = True
                        dg.RowHeadersWidth = 25

                        dg.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dg.AutoResizeRows()


                        'dg.TableStyles.Clear()
                        'Dim ts1 As DataGridTableStyle = New DataGridTableStyle
                        'ts1.MappingName = dtbl.TableName
                        'ts1.AllowSorting = False
                        'dg.TableStyles.Add(ts1)

                        'Dim myg As DataGridTableStyle
                        ''Dim myg As DataGridTableStyle
                        ''Dim myc As DataGridColumnStyle
                        'Dim myc As DataGridColumnStyle
                        'Count2 = 0
                        'For Each myc In dg.TableStyles(0).GridColumnStyles
                        '    'myc.Alignment = HorizontalAlignment.Center
                        '    myc.NullText = ""
                        '    Count2 = Count2 + 1
                        'Next
                        dg.Refresh()

                        'autosize datagrid
                        ''''debugWriteLine("2: " & Count1)

                        'Call AutoSizeGrid(500, dv, dg, dv.Count, Count2, 0, False)
                    Catch ex As Exception
                        MsgBox("Else" & ChrW(10) & ex.Message)
                    End Try

                ElseIf Count1 = 4 Then 'tblCompanyData

                    Try

                        'add another column
                        Dim tcol3 As New DataColumn
                        tcol3.DataType = System.Type.GetType("System.String")
                        tcol3.ColumnName = "Example"
                        tcol3.Caption = "Example"
                        tcol3.ReadOnly = False
                        dtbl.Columns.Add(tcol3)

                        'add another column
                        Dim tcol3a As New DataColumn
                        tcol3a.DataType = System.Type.GetType("System.String")
                        tcol3a.ColumnName = "charTab"
                        tcol3a.Caption = "Tab"
                        tcol3a.ReadOnly = False
                        dtbl.Columns.Add(tcol3a)

                        'add another column
                        Dim tcol4a As New DataColumn
                        'tcol4a.DataType = System.Type.GetType("System.String")
                        tcol4a.DataType = System.Type.GetType("System.Boolean")
                        tcol4a.ColumnName = "boolIsBool"
                        tcol4a.Caption = "boolIsBool"
                        tcol4a.ReadOnly = False
                        dtbl.Columns.Add(tcol4a)

                        'add another column
                        Dim tcol5a As New DataColumn
                        tcol5a.DataType = System.Type.GetType("System.String")
                        tcol5a.ColumnName = "boolDD"
                        tcol5a.Caption = "boolDD"
                        tcol5a.ReadOnly = False
                        dtbl.Columns.Add(tcol5a)

                        Dim strF As String
                        dgv = frmH.dgvDataCompany

                        strF = "charTab = 'Data'"
                        Dim dvA As System.Data.DataView = New DataView(dtbl, strF, "", DataViewRowState.CurrentRows)

                        dgv.DataSource = dvA
                        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                        int1 = dgv.Columns.Count
                        For Count2 = 0 To int1 - 1
                            dgv.AutoResizeColumn(Count2, DataGridViewAutoSizeColumnMode.AllCells)
                            dgv.Columns.Item(Count2).SortMode = DataGridViewColumnSortMode.NotSortable
                        Next
                        dgv.Columns.Item(0).ReadOnly = True
                        dgv.AllowUserToResizeColumns = True
                        dgv.AllowUserToResizeRows = True
                        dgv.RowHeadersWidth = 25
                        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                        Try
                            dgv.Columns.Item(1).MinimumWidth = 150
                        Catch ex As Exception
                            var1 = ex.Message
                        End Try
                        dgv.Columns.Item("Example").Visible = False
                        dgv.Columns.Item("Example").ReadOnly = True
                        dgv.Columns.Item("charTab").Visible = False
                        dgv.Columns.Item("charTab").ReadOnly = False

                        dgv.Columns.Item("boolIsBool").Visible = False
                        dgv.Columns.Item("boolIsBool").ReadOnly = False
                        dgv.Columns.Item("boolDD").Visible = False
                        dgv.Columns.Item("boolDD").ReadOnly = False

                        dgv.AutoResizeColumns()


                        strF = "charTab = 'Config'"
                        dgv = frmH.dgvStudyConfig
                        Dim dvB As System.Data.DataView = New DataView(dtbl, strF, "", DataViewRowState.CurrentRows)

                        dgv.DataSource = dvB
                        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                        int1 = dgv.Columns.Count
                        For Count2 = 0 To int1 - 1
                            dgv.AutoResizeColumn(Count2, DataGridViewAutoSizeColumnMode.AllCells)
                            dgv.Columns.Item(Count2).SortMode = DataGridViewColumnSortMode.NotSortable
                        Next
                        dgv.Columns.Item(0).ReadOnly = True
                        dgv.AllowUserToResizeColumns = True
                        dgv.AllowUserToResizeRows = True
                        dgv.RowHeadersWidth = 25
                        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                        Try
                            dgv.Columns.Item(1).MinimumWidth = 150
                        Catch ex As Exception
                            var1 = ex.Message
                        End Try

                        dgv.Columns.Item("Example").Visible = False
                        dgv.Columns.Item("Example").ReadOnly = True
                        dgv.Columns.Item("charTab").Visible = False
                        dgv.Columns.Item("charTab").ReadOnly = False

                        dgv.Columns.Item("boolIsBool").Visible = False
                        dgv.Columns.Item("boolIsBool").ReadOnly = False
                        dgv.Columns.Item("boolDD").Visible = False
                        dgv.Columns.Item("boolDD").ReadOnly = False

                        dgv.AutoResizeColumns()

                    Catch ex As Exception
                        MsgBox("Problem creating table " & Count1 & ChrW(10) & ex.Message)
                    End Try


                ElseIf Count1 = 5 Then

                    'debug
                    Try
                        int1 = dv.Count
                        'MsgBox("int1 = dv.Count: " & int1)
                    Catch ex As Exception
                        MsgBox("int1 = dv.Count:" & ChrW(10) & ex.Message)
                    End Try

                    Try
                        dgv.DataSource = dv
                    Catch ex As Exception
                        MsgBox("dgv.DataSource = dv" & ChrW(10) & ex.Message)
                    End Try
                    Try
                        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                    Catch ex As Exception
                        MsgBox("dgv.ColumnHeadersDefaultCellStyle." & ChrW(10) & ex.Message)
                    End Try

                    Try

                        Try
                            int1 = dgv.Columns.Count
                            'MsgBox("int1 = dgv.Columns.Count: " & int1)
                        Catch ex As Exception
                            MsgBox("int1 = dgv.Columns.Count" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                        Catch ex As Exception
                            MsgBox("dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
                        Catch ex As Exception
                            MsgBox("dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
                        Catch ex As Exception
                            MsgBox("dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            For Count2 = 0 To int1 - 1
                                dgv.Columns.Item(Count2).SortMode = DataGridViewColumnSortMode.NotSortable
                            Next
                        Catch ex As Exception
                            MsgBox("For Count2 = 0 To int1 - 1" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.Columns.Item(0).ReadOnly = True
                        Catch ex As Exception
                            MsgBox("dgv.Columns.Item(0).ReadOnly = True" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.AllowUserToResizeColumns = True
                        Catch ex As Exception
                            MsgBox("dgv.AllowUserToResizeColumns = True" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.AllowUserToResizeRows = True
                        Catch ex As Exception
                            MsgBox("dgv.AllowUserToResizeRows = True" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.RowHeadersWidth = 25
                        Catch ex As Exception
                            MsgBox("dgv.RowHeadersWidth = 25" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            If int1 < 2 Then
                            Else
                                var1 = dgv.Columns.Item(1).MinimumWidth 'debug
                                var2 = dgv.Columns(1).Width 'debug
                                'dgv.Columns.Item(1).MinimumWidth = 150
                            End If
                        Catch ex As Exception
                            'MsgBox("dgv.Columns.Item(1).MinimumWidth = 150" & ChrW(10) & ex.Message)
                        End Try

                        'try it again. .NET 4.6.1 thing
                        Try
                            int1 = dgv.ColumnCount
                            If int1 < 2 Then
                            Else
                                var1 = dgv.Columns.Item(1).MinimumWidth 'debug
                                var2 = dgv.Columns(1).Width 'debug
                                'dgv.Columns(1).MinimumWidth = 150
                            End If
                        Catch ex As Exception
                            ' MsgBox("dgv.Columns.Item(1).MinimumWidth = 150" & ChrW(10) & ex.Message)
                        End Try

                        Try
                            dgv.Columns(0).Width = 150
                        Catch ex As Exception
                            MsgBox(" dgv.Columns(0).Width = 150" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.AutoResizeColumns()
                        Catch ex As Exception
                            MsgBox("dgv.AutoResizeColumns()" & ChrW(10) & ex.Message)
                        End Try
                        Try
                            dgv.AutoResizeRows()
                        Catch ex As Exception
                            MsgBox(" dgv.AutoResizeRows()" & ChrW(10) & ex.Message)
                        End Try

                    Catch ex As Exception
                        MsgBox("int1 = dgv.Columns.Count" & ChrW(10) & ex.Message)
                    End Try

                Else

                    'Try
                    '    dg.DataSource = dv
                    '    'dg.DataSource = dtbl
                    '    dg.Refresh()

                    '    dg.TableStyles.Clear()
                    '    Dim ts1 As DataGridTableStyle = New DataGridTableStyle
                    '    ts1.MappingName = dtbl.TableName
                    '    ts1.AllowSorting = False
                    '    dg.TableStyles.Add(ts1)

                    '    Dim myg As DataGridTableStyle
                    '    'Dim myg As DataGridTableStyle
                    '    'Dim myc As DataGridColumnStyle
                    '    Dim myc As DataGridColumnStyle
                    '    Count2 = 0
                    '    For Each myc In dg.TableStyles(0).GridColumnStyles
                    '        'myc.Alignment = HorizontalAlignment.Center
                    '        myc.NullText = ""
                    '        Count2 = Count2 + 1
                    '    Next
                    '    dg.Refresh()

                    '    'autosize datagrid
                    '    ''''debugWriteLine("2: " & Count1)

                    '    Call AutoSizeGrid(500, dv, dg, dv.Count, Count2, 0, False)
                    'Catch ex As Exception
                    '    MsgBox("Else" & ChrW(10) & ex.Message)
                    'End Try




                    ''rsdt.Filter = ""
                End If
            Next
        Catch ex As Exception
            MsgBox("Count1: " & Count1 & ChrW(10) & ex.Message)
        End Try


        'rsdt.Close()
        'configure tblReportTableConfig
        Try
            Call FillTableReports(True)
        Catch ex As Exception
            MsgBox("Call FillTableReports(True)" & ChrW(10) & ex.Message)
        End Try

        'pesky
        Try
            Call OrderReportTableConfig()
        Catch ex As Exception
            MsgBox(" Call OrderReportTableConfig()" & ChrW(10) & ex.Message)
        End Try


        'prepare tbl_dgHome
        'dgHome does not contain any row titles
        Try
            Call Prepare_tbl(tbl_dgHome, frmH.dgHome)
        Catch ex As Exception
            MsgBox("Call Prepare_tbl(tbl_dgHome, frmH.dgHome)" & ChrW(10) & ex.Message)
        End Try

        frmH.dgHome.DataSource = tbl_dgHome
        ''''debugWriteLine("3")


        'Call AutoSizeGrid(500, tbl_dgHome, frmh.dgHome, tbl_dgHome.Rows.Count, tbl_dgHome.Columns.Count, 0, False)

        'dgCompanyAnalRef and dgWatsonAnalRef 1st columns must be synchronized
        wid1 = 1
        wid2 = 1
        'synchronize col widths
        frmH.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
        Call SyncCols(frmH.dgvWatsonAnalRef, frmH.dgvCompanyAnalRef)

        ''make dgDataWatson Data tables column2width 150
        'Try
        '    frmH.dgDataWatson.TableStyles(0).GridColumnStyles(1).Width = 150
        'Catch ex As Exception
        '    MsgBox("frmH.dgDataWatson.TableStyles(0).GridColumnStyles(1).Width = 150" & ChrW(10) & ex.Message)
        'End Try


        'prepare Analytical Run Summary
        dtbl = tblAnalRunSum
        'Dim ts2 As New DataGridTableStyle
        Try
            For Count1 = 0 To 20
                Select Case Count1
                    Case 0
                        str1 = "boolInclude"
                        str2 = "A *"
                        boolRO = False
                    Case 1
                        str1 = "boolIncludeRegr"
                        str2 = "B *"
                        boolRO = False
                    Case 2
                        str1 = "Watson Run ID"
                        str2 = "Watson" & Chr(13) & "Run ID"
                        boolRO = True
                    Case 3
                        str1 = "Analyte"
                        str2 = str1
                        boolRO = True

                        '20160205 LEE: Added column
                    Case 4
                        str1 = "Analyte_C"
                        str2 = str1
                        boolRO = True

                        '20160205 LEE: Added column
                    Case 5
                        str1 = "Matrix"
                        str2 = str1
                        boolRO = True

                        '20160319 LEE: Added column
                    Case 6
                        str1 = "Run Type"
                        str2 = str1
                        boolRO = True

                    Case 7
                        str1 = "Notebook ID"
                        str2 = "Notebook ID"
                        boolRO = True

                    Case 8
                        '20171108 LEE: Added for Alturas
                        str1 = "Instrument ID"
                        str2 = "Instrument ID"
                        boolRO = True 'will set the gridstyle readonly as True

                    Case 9
                        str1 = "Extraction Date"
                        str2 = "Extraction Date"
                        boolRO = True
                    Case 10
                        str1 = "Analysis Date"
                        str2 = "Analysis Date"
                        boolRO = True
                    Case 11
                        str1 = "Pass/Fail"
                        str2 = "Pass/Fail"
                        boolRO = True
                    Case 12
                        str1 = "Samples"
                        str2 = "Run Description" 'str1
                        boolRO = True
                    Case 13
                        str1 = "Watson Comments"
                        str2 = "Watson Comments"
                        boolRO = True
                    Case 14
                        str1 = "User Comments"
                        str2 = "User Comments"
                        boolRO = False 'will set the gridstyle readonly as false
                    Case 15
                        str1 = "RUNTYPEID"
                        str2 = str1
                        boolRO = True 'will set the gridstyle readonly as True
                    Case 16
                        str1 = "LLOQ"
                        str2 = str1
                        boolRO = True 'will set the gridstyle readonly as True
                    Case 17
                        str1 = "ULOQ"  ''
                        str2 = str1
                        boolRO = True 'will set the gridstyle readonly as True

                        'NDL 20-Jan-2016 Added 2 extra columns (these should not be visible)
                    Case 18
                        str1 = "AnalyteID"  ''
                        str2 = str1
                        boolRO = True 'will set the gridstyle readonly as True
                    Case 19
                        str1 = "boolInThisRunsAssayID"  ''
                        str2 = str1
                        boolRO = True 'will set the gridstyle readonly as True
                    Case 20
                        str1 = "RUNANALYTEREGRESSIONSTATUS"  ''
                        str2 = str1
                        boolRO = True 'will set the gridstyle readonly as True



                End Select
                Dim dc As New DataColumn
                Select Case Count1
                    Case 20
                        dc.DataType = System.Type.GetType("System.Int16")
                    Case Else
                        If Count1 < 2 Then
                            dc.DataType = System.Type.GetType("System.Boolean")
                        Else
                            dc.DataType = System.Type.GetType("System.String")
                        End If
                End Select

                dc.ColumnName = str1
                dc.Caption = str2
                dc.ReadOnly = boolRO
                dc.AllowDBNull = True

                dtbl.Columns.Add(dc)

            Next
        Catch ex As Exception
            MsgBox("For Count1 = 0 To 18: " & Count1 & ChrW(10) & ex.Message)
        End Try

        dv = dtbl.DefaultView
        'some change for vss testing
        dv.AllowNew = False
        dv.AllowDelete = False
        dv.AllowEdit = True

        Try
            Dim dgv1 As DataGridView = frmH.dgvAnalyticalRunSummary

            dgv1.DataSource = dv
            dgv1.Columns.Item("Analyte").Frozen = True
            dgv1.Columns.Item("Analyte_C").Frozen = True
            int1 = dgv1.Columns.Count
            For Count1 = 0 To int1 - 1
                dgv1.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv1.Columns.Item(Count1).ReadOnly = True
            Next
            dgv1.Columns.Item("boolInclude").HeaderText = "A *"
            dgv1.Columns.Item("boolIncludeRegr").HeaderText = "B *"
            dgv1.Columns.Item("Watson Run ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgv1.AllowUserToResizeColumns = True
            dgv1.AllowUserToResizeRows = True
            dgv1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
            dgv1.Columns.Item("Watson Run ID").Width = 60
            dgv1.Columns.Item("User Comments").Width = 500
            dgv1.RowHeadersWidth = 25
            dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv1.Columns.Item("RUNTYPEID").Visible = False

            dgv1.Columns.Item("Samples").HeaderText = "Run Description"

            'NDL 20-Jan-2016 These are not for display
            dgv1.Columns.Item("AnalyteID").Visible = False
            dgv1.Columns.Item("boolInThisRunsAssayID").Visible = False

            dgv1.Columns("Instrument ID").HeaderText = "Instrument ID"

        Catch ex As Exception
            MsgBox("Dim dgv1 As DataGridView = frmH.dgvAnalyticalRunSummary" & ChrW(10) & ex.Message)
        End Try


        ''rsdt = Nothing
        dtbl = Nothing

        'MsgBox("End of CreateTables")

    End Sub

    Sub MethValAutoCol()

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim var1
        Dim num1 As Single
        Dim num2 As Single
        Dim num3 As Single
        Dim num4 As Single
        Dim num5 As Single
        Dim int1 As Short

        dgv = frmH.dgvMethodValData

        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()

        num1 = dgv.Width
        num2 = dgv.Columns(0).Width
        int1 = dgv.Columns.Count
        num3 = num1 - num2
        num4 = (num3 / (int1 - 1)) * 0.8

        For Count1 = 1 To int1 - 1
            dgv.Columns(Count1).MinimumWidth = num4
            dgv.Columns(Count1).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        Next

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        'dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()


        dgv.Columns(0).Width = num2



    End Sub

    Sub SetAnalRunSummaryColumns()

        Dim Count1 As Short
        Dim int1 As Short

        Dim dgv1 As DataGridView = frmH.dgvAnalyticalRunSummary

        If dgv1.Columns.Item("RUNTYPEID").Visible = True Then
            dgv1.Columns.Item("Analyte").Frozen = True
            int1 = dgv1.Columns.Count
            For Count1 = 0 To int1 - 1
                dgv1.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv1.Columns.Item(Count1).ReadOnly = True
            Next
            dgv1.Columns.Item("boolInclude").HeaderText = "A *"
            dgv1.Columns.Item("boolIncludeRegr").HeaderText = "B *"
            dgv1.Columns.Item("Watson Run ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgv1.AllowUserToResizeColumns = True
            dgv1.AllowUserToResizeRows = True
            dgv1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
            dgv1.Columns.Item("Watson Run ID").Width = 60
            dgv1.Columns.Item("User Comments").Width = 500
            dgv1.RowHeadersWidth = 25
            dgv1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv1.Columns.Item("RUNTYPEID").Visible = False
            'NDL 20-Jan-2016 These are not for display
            dgv1.Columns.Item("ANALYTEID").Visible = False
            dgv1.Columns.Item("boolInThisRunsAssayID").Visible = False

        End If

        '20160205 LEE: configure new column
        If boolAnalyte_C() Then
            dgv1.Columns.Item("Analyte_C").Visible = True
        Else
            dgv1.Columns.Item("Analyte_C").Visible = False
        End If
        'dgv1.Columns.Item("Analyte_C").Visible = True
        dgv1.Columns.Item("Analyte_C").HeaderText = "Analyte C[n]"

        'boolMultiMatrix
        If boolMultiMatrix() Then
            dgv1.Columns.Item("Matrix").Visible = True
        Else
            dgv1.Columns.Item("Matrix").Visible = False
        End If
        dgv1.Columns.Item("Matrix").HeaderText = "Matrix"

    End Sub

    Sub SyncCols(ByVal dgv1 As DataGridView, ByVal dgv2 As DataGridView)

        'dgv1=Watson, dgv2=Company
        Dim int1 As Short
        Dim Count1 As Short
        Dim wid1, wid2
        Dim str1 As String
        'W=1
        'C=2
        int1 = dgv2.Columns.Count

        'first do column(Item)
        wid1 = dgv2.Columns.Item("BOOLINCLUDE").Width + dgv2.Columns.Item("Item").Width
        dgv1.Columns.Item("Item").Width = wid1

        For Count1 = 0 To int1 - 1
            str1 = dgv2.Columns.Item(Count1).Name
            If StrComp(str1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'IGNORE
            ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'IGNORE
            ElseIf StrComp(str1, "Item", CompareMethod.Text) = 0 Then 'IGNORE
            Else
                If dgv1.Columns.Contains(str1) Then
                    wid1 = dgv2.Columns.Item(str1).Width
                    wid2 = dgv1.Columns.Item(str1).Width
                    If wid1 > wid2 Then
                        dgv1.Columns.Item(str1).Width = wid1
                    ElseIf wid2 > wid1 Then
                        dgv2.Columns.Item(str1).Width = wid2
                    End If
                End If
            End If
        Next

    End Sub

    Sub InitializeSummaryData()

        Dim tbl1 As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim var1, var2
        Dim intCols As Short
        Dim strM As String = ""

        Try
            strF = "id_tblStudies = -1"
            strS = "intOrder ASC"

            Try
                tbl1 = tblSummaryData
            Catch ex As Exception
                MsgBox("Sub InitializeSummaryData: tbl1 = tblSummaryData:" & ChrW(10) & ex.Message)
            End Try


            'add unbound column to tbl1
            Dim col1 As New DataColumn
            col1.ColumnName = "boolI"
            col1.Caption = "A *"
            col1.DataType = System.Type.GetType("System.Boolean")
            tbl1.Columns.Add(col1)

            Try
                dgv = frmH.dgvSummaryData
            Catch ex As Exception
                MsgBox("Sub InitializeSummaryData: dgv = frmH.dgvSummaryData:" & ChrW(10) & ex.Message)
            End Try

            dgv.RowHeadersWidth = 25
            dgv.AllowUserToOrderColumns = False
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill)
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

            Dim dv As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
            dv.AllowDelete = False
            dv.AllowNew = False
            dgv.DataSource = dv

            dgv.AutoResizeColumns()

            int1 = dgv.Columns.Count

            'MsgBox("int1 = dgv.Columns.Count: " & int1)

            For Count1 = 0 To int1 - 1
                dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DisplayIndex = int1 - 1
            Next

            If int1 < 1 Then
            Else

                Call OrderSummaryTable()

                dgv.Columns.Item("boolI").Visible = True
                dgv.Columns.Item("boolI").HeaderText = "A *"
                dgv.Columns.Item("intOrder").Visible = True
                dgv.Columns.Item("intOrder").HeaderText = "Order"
                dgv.Columns.Item("intOrder").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv.Columns.Item("charRowName").Visible = True
                dgv.Columns.Item("charRowName").HeaderText = "Item"
                Try
                    dgv.Columns.Item("charRowName").MinimumWidth = 200
                Catch ex As Exception
                    'MsgBox("Sub InitializeSummaryData: dgv.Columns.Item('charRowName').MinimumWidth = 200: " & ChrW(10) & ex.Message)
                    strM = "Sub InitializeSummaryData: dgv.Columns.Item('charRowName').MinimumWidth = 200: " & ChrW(10) & ex.Message
                End Try
                dgv.Columns.Item("charValue").Visible = True
                dgv.Columns.Item("charValue").HeaderText = "Value"
                dgv.Columns.Item("charValue").MinimumWidth = 150
                dgv.Columns.Item("charvalue").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                'calculate min width
                var1 = dgv.RowHeadersWidth + dgv.Columns.Item("charRowName").Width + dgv.Columns.Item("boolI").Width + dgv.Columns.Item("intOrder").Width
                var2 = dgv.Width

                Try
                    dgv.Columns.Item("charValue").MinimumWidth = (var2 - var1) * 0.97
                Catch ex As Exception
                    strM = "Sub InitializeSummaryData: dgv.Columns.Item('charValue').MinimumWidth = (var2 - var1) * 0.97: " & ChrW(10) & ex.Message
                End Try
                'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

                dgv.Columns.Item("charRowName").ReadOnly = True
                dgv.Columns.Item("charValue").ReadOnly = True
            End If

        Catch ex As Exception
            MsgBox("Sub InitializeSummaryData:" & ChrW(10) & ex.Message)
        End Try


    End Sub



    Sub OrderSummaryTable()

        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short

        dgv = frmH.dgvSummaryData
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item("boolI").DisplayIndex = 0
            dgv.Columns.Item("intOrder").DisplayIndex = 1
            dgv.Columns.Item("charRowName").DisplayIndex = 2
            dgv.Columns.Item("charValue").DisplayIndex = 3

        Next

        'position cmdOrder
        Dim wd1, wd2
        wd1 = dgv.Left
        wd1 = wd1 + dgv.RowHeadersWidth
        wd1 = wd1 + dgv.Columns.Item("boolI").Width
        wd2 = dgv.Columns.Item("intOrder").Width
        frmH.cmdOrderSummaryTable.Left = wd1
        'frmH.cmdOrderSummaryTable.Width = wd2 * 1.1

    End Sub

    Sub ConfigureSummaryTable()

        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim ctRows1 As Short
        Dim ctRows2 As Short
        Dim strF As String
        Dim strS As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim var1, var2
        Dim boolUpdate As Boolean

        boolUpdate = False

        tbl1 = tblSummaryData
        tbl2 = tblDataTableRowTitles
        strF = "id_tblStudies = " & id_tblStudies
        strS = "intOrder ASC"
        'first check to see if data exists
        rows1 = tbl1.Select(strF, strS)
        ctRows1 = rows1.Length

        strF = "charDataTableName = 'tblSummaryData' and boolInclude = -1" ' & True
        strS = "intOrder ASC"
        rows2 = tbl2.Select(strF, strS, DataViewRowState.CurrentRows)
        ctRows2 = rows2.Length

        dgv = frmH.dgvSummaryData

        'check to see if charRowName is updated
        For Count1 = 0 To ctRows1 - 1
            var1 = rows1(Count1).Item("id_tblDataTableRowTitles")
            strF = "id_tblDataTableRowTitles = " & var1
            rows2 = tbl2.Select(strF)
            int1 = rows2.Length
            If int1 <> 0 Then 'see if modification is needed
                var2 = rows2(0).Item("CHARROWNAME")
                var1 = rows1(Count1).Item("CHARROWNAME")
                If StrComp(CStr(var1), CStr(var2), CompareMethod.Text) = 0 Then 'ignore
                Else
                    rows1(Count1).BeginEdit()
                    rows1(Count1).Item("CHARROWNAME") = var2
                    rows1(Count1).EndEdit()
                    boolUpdate = True
                End If
            End If
        Next

        If ctRows1 = 0 Then 'add rows to table
            For Count1 = 0 To ctRows2 - 1
                Dim nr As DataRow = tbl1.NewRow
                nr.BeginEdit()
                nr("id_tblStudies") = id_tblStudies
                nr("id_tblDataTableRowTitles") = rows2(Count1).Item("id_tblDataTableRowTitles")
                nr("charRowName") = rows2(Count1).Item("charRowName")
                nr("charValue") = ""
                nr("boolInclude") = -1 'True
                nr("boolI") = True
                nr("intOrder") = Count1 + 1
                nr.EndEdit()
                tbl1.Rows.Add(nr)
            Next

            'reset rows1
            strF = "id_tblStudies = " & id_tblStudies
            Erase rows1
            rows1 = tbl1.Select(strF)
            ctRows1 = rows1.Length

            boolUpdate = True

        End If

        strF = "charDataTableName = 'tblSummaryData' and boolInclude = -1" ' & True
        strS = "intOrder ASC"
        rows2 = tbl2.Select(strF, strS, DataViewRowState.CurrentRows)
        ctRows2 = rows2.Length

        'now check to see if more default items have been added
        If ctRows2 > ctRows1 Then 'add more rows
            Erase rows1
            For Count1 = 0 To ctRows2 - 1
                'var1 = rows2(Count1).Item("id_tblDataTableRowTitles")

                int1 = rows2.Length 'debug
                var1 = rows2(Count1).Item("id_tblDataTableRowTitles")


                strF = "id_tblDataTableRowTitles = " & var1 & " AND id_tblStudies = " & id_tblStudies
                rows1 = tbl1.Select(strF)
                int1 = rows1.Length
                If int1 = 0 Then 'add row
                    Dim nr As DataRow = tbl1.NewRow
                    nr.BeginEdit()
                    nr("id_tblStudies") = id_tblStudies
                    nr("id_tblDataTableRowTitles") = rows2(Count1).Item("id_tblDataTableRowTitles")
                    nr("charRowName") = rows2(Count1).Item("charRowName")
                    nr("charValue") = ""
                    nr("boolInclude") = -1
                    nr("boolI") = True
                    nr("intOrder") = rows2(Count1).Item("intOrder")
                    nr.EndEdit()
                    tbl1.Rows.Add(nr)
                End If
            Next
            boolUpdate = True
        End If

        If boolUpdate Then
            'frmh.ta_tblSummaryData.Update(tblSummaryData)
            'frmh.ta_tblSummaryData.Fill(tblSummaryData)

            If boolGuWuOracle Then
                Try
                    ta_tblSummaryData.Update(tblSummaryData)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLSUMMARYDATA.Merge('ds2005.TBLSUMMARYDATA, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblSummaryDataAcc.Update(tblSummaryData)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLSUMMARYDATA.Merge('ds2005Acc.TBLSUMMARYDATA, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblSummaryDataSQLServer.Update(tblSummaryData)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLSUMMARYDATA.Merge('ds2005Acc.TBLSUMMARYDATA, True)
                End Try
            End If

        End If

        'update boolI in tbl
        Call UpdateBoolSummaryTable()

        'update charValue column
        Try
            Call UpdateValueSummaryTable()
        Catch ex As Exception

        End Try


        Dim bool As Boolean
        If frmH.rbShowIncludedSummaryTable.Checked = True Then
            strF = "id_tblStudies = " & id_tblStudies & " AND boolInclude = -1"
        Else
            strF = "id_tblStudies = " & id_tblStudies
        End If
        strS = "intOrder ASC"
        Dim dv As System.Data.DataView = New DataView(tbl1, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowNew = False
        dgv.DataSource = dv

        'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
        'dgv.AutoResizeColumns()

    End Sub

    Function GetSummaryString(ByVal dv As System.Data.DataView, ByVal strRow As String, ByVal arrM As Array, ByVal arrM1 As Array, ByVal intMeth1 As Short, ByVal intMeth As Short, ByVal intM As Short, ByVal boolListAll As Boolean) As String


        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2, var3, var4
        Dim varE
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim intHit As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String


        GetSummaryString = ""

        'dv = frmH.dgvMethodValData.DataSource
        int2 = FindRowDV(strRow, dv)
        If intMeth1 = 1 And boolListAll = False Then
            var3 = arrM(1, 1)
            var2 = NZ(dv(int2).Item(arrM(2, 1)), "[None]")
            GetSummaryString = var2
        Else
            str1 = ""
            For Count2 = 1 To intMeth1
                var1 = arrM1(intM, Count2)
                var3 = NZ(dv(int2).Item(arrM1(2, Count2)), "[None]")
                str1 = str1 & """" & var3 & """"
                intHit = 0
                For Count3 = 1 To intMeth
                    var2 = arrM(intM, Count3)
                    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                        intHit = intHit + 1
                        If intHit = 1 Then
                            var3 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                            str1 = str1 & " for " & arrM(1, Count3)
                        Else
                            str1 = str1 & " and " & arrM(1, Count3)
                        End If
                    End If
                Next
                str1 = str1 & ChrW(10)
                If intHit = intMeth1 Then
                    Exit For
                End If
            Next
            'remove last line return
            GetSummaryString = Mid(str1, 1, Len(str1) - 1)
        End If

    End Function

    Sub UpdateValueSummaryTable()

        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim tblD As System.Data.DataTable
        Dim rowsD() As DataRow
        Dim strF As String
        Dim strS As String
        Dim int1 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim var1, var2, var3, var4
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim int5 As Short
        Dim int6 As Short
        Dim int7 As Short
        Dim int8 As Short
        Dim int9 As Short
        Dim int10 As Short
        Dim int11 As Short
        Dim int12 As Short
        Dim dv As System.Data.DataView
        Dim varE
        Dim arrM(10, 50)
        '1=cmpd name, 2=position, 3=Val Title, 4=Lab Method, 5=Val Protocol #, 6=Val Study #
        Dim arrM1(10, 50)
        '1=LabMethod, 2=position, 3=Val Title, 4=Lab Method, 5=Val Protocol #, 6=Val Study #
        Dim intMeth As Short
        Dim intMeth1 As Short
        Dim dg As DataGrid
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim varR
        Dim intCt As Short
        Dim dgv As DataGridView
        Dim rowsX() As DataRow
        Dim boolGo As Boolean
        Dim intCols As Short
        Dim intHit As Short
        Dim intMS As Short
        Dim boolListAll As Boolean
        Dim boolVal As Boolean
        Dim idR As Int64
        Dim intA(1) As Short
        Dim dgv1 As DataGridView
        Dim strUnits As String

        Dim dgvR As DataGridView
        dgvR = frmH.dgvReports
        boolVal = False
        If dgvR.Rows.Count = 0 Then
            boolVal = False
        Else
            idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
            If idR > 1 And idR < 1000 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If

        dgv1 = frmH.dgvMethodValData

        'first determine relevent columns in Method Validation Tab
        'int4 = dgv1.TableStyles(0).GridColumnStyles.Count
        intCols = dgv1.Columns.Count
        int1 = 1
        If intCols = 2 Then
            'arrM(1, 1) = dgv1.TableStyles(0).GridColumnStyles(1).MappingName
            arrM(1, 1) = dgv1.Columns(1).HeaderText
            arrM(2, 1) = 1
        Else

        End If

        intMS = 13 'Remember to update this if adding items!!!!
        ReDim arrM(intMS, 50)
        ReDim arrM1(intMS, 50)
        ReDim intA(intMS)

        dv = dgv1.DataSource
        intA(3) = FindRowDV("Validation Report Title", dv)
        intA(4) = FindRowDV("Lab Method Title", dv)
        intA(5) = FindRowDV("Validation Protocol Number", dv)
        intA(6) = FindRowDV("Validation Corporate Study/Project Number", dv)
        intA(7) = FindRowDV("Validation Report Number", dv)
        intA(8) = FindRowDV("Extraction Procedure Description", dv)
        intA(9) = FindRowDV("Species", dv)
        intA(10) = FindRowDV("Matrix", dv)
        intA(11) = FindRowDV("Anticoagulant/Preservative", dv)
        intA(12) = FindRowDV("Sample Size", dv)
        intA(13) = FindRowDV("Lab Method Number", dv)


        ''arrM(1, 1) = dgv1.TableStyles(0).GridColumnStyles(1).MappingName
        'arrM(1, 1) = dgv1.Columns(1).HeaderText
        'arrM(2, 1) = 1
        'arrM(3, 1) = NZ(dv(int3).Item(1), "NA")
        'arrM(4, 1) = NZ(dv(int4).Item(1), "NA")
        'arrM(5, 1) = NZ(dv(int5).Item(1), "NA")
        'arrM(6, 1) = NZ(dv(int6).Item(1), "NA")
        int1 = 0

        Dim intSSS As Short
        For Count1 = 1 To intCols - 1
            'var1 = dgv1.TableStyles(0).GridColumnStyles(Count2).MappingName
            var2 = dgv1.Columns(Count1).HeaderText

            int1 = int1 + 1
            arrM(1, int1) = var2
            arrM(2, int1) = Count1
            For Count2 = 3 To intMS
                'Select Case Count2
                '    Case 3
                '        intSSS = 3
                '    Case 4
                '        intSSS = 4
                '    Case 5
                '        intSSS = 5
                '    Case 6
                '        intSSS = 6
                '    Case 7
                '        intSSS = 7
                '    Case 8
                '        intSSS = 8
                '    Case 9
                '        intSSS = 9
                '    Case 10
                '        intSSS = 10
                '    Case 11
                '        intSSS = 11
                '    Case 12
                '        intSSS = 12

                'End Select

                var1 = dv(intA(Count2)).Item(Count1)
                arrM(Count2, int1) = NZ(dv(intA(Count2)).Item(Count1), "NA")
            Next
            'arrM(3, int1) = NZ(dv(int3).Item(Count1), "NA")
            'arrM(4, int1) = NZ(dv(int4).Item(Count1), "NA")
            'arrM(5, int1) = NZ(dv(int5).Item(Count1), "NA")
            'arrM(6, int1) = NZ(dv(int6).Item(Count1), "NA")
            'arrM(7, int1) = NZ(dv(int7).Item(Count1), "NA")
            'arrM(8, int1) = NZ(dv(int8).Item(Count1), "NA")

        Next
        intMeth = int1

        'now check for unique lab methods
        dv = dgv1.DataSource
        int2 = FindRowDV("Lab Method Title", dv)
        intMeth1 = 0
        If intMeth > 1 Then
            intMeth1 = 1
            arrM1(1, 1) = arrM(4, 1)
            arrM1(2, 1) = intMeth1
            For Count3 = 3 To intMS
                var1 = arrM(Count3, 1)
                arrM1(Count3, intMeth1) = arrM(Count3, 1)
            Next
            'arrM1(3, intMeth1) = arrM(3, 1)
            'arrM1(4, intMeth1) = arrM(4, 1)
            'arrM1(5, intMeth1) = arrM(5, 1)
            'arrM1(6, intMeth1) = arrM(6, 1)
            'arrM1(7, intMeth1) = arrM(7, 1)
            For Count1 = 1 To intMeth
                If Count1 = 9 Then
                    var1 = "a"
                End If
                var1 = arrM(4, Count1)
                For Count2 = Count1 + 1 To intMeth
                    var2 = arrM(4, Count2)
                    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                    Else
                        'ensure is unique
                        If intMeth1 = 0 Then
                            intMeth1 = intMeth1 + 1
                            arrM1(1, intMeth1) = var2
                            arrM1(2, intMeth1) = Count2
                            For Count3 = 3 To intMS
                                arrM1(Count3, intMeth1) = arrM(Count3, Count2)
                            Next
                            'arrM1(3, intMeth1) = arrM(3, Count2)
                            'arrM1(4, intMeth1) = arrM(4, Count2)
                            'arrM1(5, intMeth1) = arrM(5, Count2)
                            'arrM1(6, intMeth1) = arrM(6, Count2)
                            'arrM1(7, intMeth1) = arrM(7, Count2)
                        Else
                            boolGo = True
                            For Count3 = 1 To intMeth1
                                var3 = arrM1(1, Count3)
                                If StrComp(var2, var3, CompareMethod.Text) = 0 Then
                                    boolGo = False
                                    Exit For
                                End If
                            Next
                            If boolGo Then
                                intMeth1 = intMeth1 + 1
                                arrM1(1, intMeth1) = var2
                                arrM1(2, intMeth1) = Count2
                                For Count3 = 3 To intMS
                                    arrM1(Count3, intMeth1) = arrM(Count3, Count2)
                                Next
                                'arrM1(3, intMeth1) = arrM(3, Count2)
                                'arrM1(4, intMeth1) = arrM(4, Count2)
                                'arrM1(5, intMeth1) = arrM(5, Count2)
                                'arrM1(6, intMeth1) = arrM(6, Count2)
                                'arrM1(7, intMeth1) = arrM(7, Count2)

                            End If
                        End If
                    End If
                Next
            Next

        Else
            intMeth1 = 1
            arrM1(1, 1) = arrM(4, 1)
            arrM1(2, 1) = intMeth1
            For Count3 = 3 To intMS
                arrM1(Count3, intMeth1) = arrM(Count3, 1)
            Next
        End If

        strF = "id_tblStudies = " & id_tblStudies
        tbl = tblSummaryData
        rows = tbl.Select(strF)
        'tblD = tblDataTableRowTitles
        intCt = rows.Length
        For Count1 = 0 To intCt - 1
            rows(Count1).BeginEdit()
            varR = NZ(rows(Count1).Item("charRowName"), "")
            'strF = "CHARDATATABLENAME = 'TBLSUMMARYDATA' AND CHARROWNAME = '" & var1 & "'"
            'Erase rowsD
            'rowsD = tblD.Select(strF)
            varE = ""
            boolListAll = True
            If Len(varR) = 0 Then
            Else
                Select Case varR
                    Case "Validation Protocol Number"
                        boolListAll = False
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 5, boolListAll)

                        'int2 = FindRowDV("Validation Protocol Number", dv)
                        'var2 = NZ(dv(int2).Item(arrM(2, 1)), "")
                        'int3 = FindRowDV("Validation Corporate Study/Project Number", dv)
                        'var3 = NZ(dv(int3).Item(arrM(2, 1)), "")
                        'If Len(var2) = 0 Then 'use var3
                        '    varE = var3
                        'Else
                        '    varE = var2
                        'End If
                    Case "Validation Report Title"
                        'var1 = NZ(frmH.lblReportTitle.Text, "")
                        'varE = var3
                        boolListAll = False
                        dv = dgv1.DataSource
                        str1 = "Validation Report Title"
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 3, boolListAll)

                        'int2 = FindRowDV("Validation Report Title", dv)
                        'If intMeth1 = 1 Then
                        '    var3 = arrM(1, 1)
                        '    var2 = NZ(dv(int2).Item(arrM(2, 1)), "")
                        '    varE = var2
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth1
                        '        var1 = arrM1(3, Count2)
                        '        var3 = NZ(dv(int2).Item(arrM1(2, Count2)), "")
                        '        str1 = str1 & """" & var3 & """"
                        '        intHit = 0
                        '        For Count3 = 1 To intMeth
                        '            var2 = arrM(3, Count3)
                        '            If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                        '                intHit = intHit + 1
                        '                If intHit = 1 Then
                        '                    var3 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                        '                    str1 = str1 & " for " & arrM(1, Count3)
                        '                Else
                        '                    str1 = str1 & " and " & arrM(1, Count3)
                        '                End If
                        '            End If
                        '        Next
                        '        str1 = str1 & ChrW(10)
                        '        If intHit = intMeth1 Then
                        '            Exit For
                        '        End If
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If

                    Case "Validation References"
                        dv = dgv1.DataSource
                        int2 = FindRowDV("Validation Protocol Number", dv)
                        int3 = FindRowDV("Validation Corporate Study/Project Number", dv)
                        If intMeth = 1 Then
                            var3 = arrM(1, 1)
                            var2 = NZ(dv(int2).Item(arrM(1, 1)), "")
                            varE = var2
                        Else
                            str1 = ""
                            For Count2 = 1 To intMeth
                                var2 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                                var3 = NZ(dv(int3).Item(arrM(2, Count2)), "")
                                If Len(var2) = 0 Then 'use var3
                                    str1 = str1 & var3 & " for " & arrM(1, Count2) & Chr(10)
                                Else
                                    str1 = str1 & var2 & " for " & arrM(1, Count2) & Chr(10)
                                End If
                            Next
                            'remove last line return
                            varE = Mid(str1, 1, Len(str1) - 1)
                        End If

                    Case "Lab Method Number"
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 4, False)

                    Case "Lab Method Title"
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 4, False)

                        'int2 = FindRowDV("Lab Method Title", dv)
                        'If intMeth = 1 Then
                        '    var2 = NZ(dv(int2).Item(arrM(1, 1)), "")
                        '    varE = var2
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth
                        '        var2 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                        '        str1 = str1 & var2 & " for " & arrM(1, Count2) & Chr(10)
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If

                    Case "Validation Corporate Study/Project Number"
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 6, False)

                        'int2 = FindRowDV("Validation Report Title", dv)
                        'If intMeth = 1 Then
                        '    var2 = NZ(dv(int2).Item(arrM(2, 1)), "")
                        '    varE = var2
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth
                        '        var2 = NZ(dv(int2).Item(arrM(1, Count2)), "")
                        '        str1 = str1 & var2 & " for " & arrM(1, Count2) & Chr(10)
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If

                    Case "Validation Report Number"
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 7, False)


                    Case "Species"
                        boolListAll = False
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 9, boolListAll)

                        'dv = dgv1.DataSource
                        'int2 = FindRowDV("Species", dv)
                        'If intMeth = 1 Then
                        '    var2 = NZ(dv(int2).Item(arrM(1, 1)), "")
                        '    If Len(var2) = 0 Then
                        '        var3 = var2
                        '    Else
                        '        var3 = Capit(var2)
                        '    End If
                        '    varE = var3
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth
                        '        var2 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                        '        If Len(var2) = 0 Then
                        '            var3 = var2
                        '        Else
                        '            var3 = Capit(var2)
                        '        End If
                        '        str1 = str1 & var3 & " for " & arrM(1, Count2) & Chr(10)
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If
                    Case "Matrix"
                        boolListAll = False
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 10, boolListAll)


                        'dv = dgv1.DataSource
                        'int2 = FindRowDV("Matrix", dv)
                        'If intMeth = 1 Then
                        '    var2 = NZ(dv(int2).Item(arrM(1, 1)), "")
                        '    If Len(var2) = 0 Then
                        '        var3 = LowerCase(var2)
                        '    Else
                        '        var3 = Capit(LowerCase(var2))
                        '    End If
                        '    varE = var3
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth
                        '        var2 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                        '        If Len(var2) = 0 Then
                        '            var3 = LowerCase(var2)
                        '        Else
                        '            var3 = Capit(LowerCase(var2))
                        '        End If
                        '        str1 = str1 & var3 & " for " & arrM(1, Count2) & Chr(10)
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If
                    Case "Anticoagulant/Preservative"
                        boolListAll = False
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 11, boolListAll)

                        'dv = dgv1.DataSource
                        'int2 = FindRowDV("Anticoagulant/Preservative", dv)
                        'If intMeth = 1 Then
                        '    var2 = NZ(dv(int2).Item(arrM(1, 1)), "")
                        '    If Len(var2) = 0 Then
                        '        var3 = var2
                        '    Else
                        '        ''find anticoagulant
                        '        'Erase rowsX
                        '        'strF = "ID_TBLDROPDOWNBOXCONTENT = " & var2
                        '        'rowsX = tblDropdownBoxContent.Select(strF)
                        '        'var3 = rowsX(0).Item("CHARVALUE")
                        '        var3 = var2
                        '        var3 = Capit(var3)
                        '    End If
                        '    varE = var3
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth
                        '        var2 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                        '        If Len(var2) = 0 Then
                        '            var3 = var2
                        '        Else
                        '            ''find anticoagulant
                        '            'Erase rowsX
                        '            'strF = "ID_TBLDROPDOWNBOXCONTENT = " & var2
                        '            'rowsX = tblDropdownBoxContent.Select(strF)
                        '            'var3 = rowsX(0).Item("CHARVALUE")
                        '            var3 = var2
                        '            var3 = Capit(var3)

                        '        End If
                        '        str1 = str1 & var3 & " for " & arrM(1, Count2) & Chr(10)
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If

                    Case "Extraction Procedure Description"

                        boolListAll = False
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 8, boolListAll)


                        'dv = dgv1.DataSource
                        ''str1 = "Extraction Procedure Description"
                        ''varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 8)

                        'int2 = FindRowDV("Extraction Procedure Description", dv)
                        'If intMeth = 1 Then
                        '    var2 = NZ(dv(int2).Item(arrM(1, 1)), "")
                        '    var3 = Capit(var2)
                        '    varE = var3
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth
                        '        var2 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                        '        var3 = Capit(var2)
                        '        str1 = str1 & var3 & " for " & arrM(1, Count2) & Chr(10)
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If
                    Case "Sample Size"
                        boolListAll = False
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryString(dv, str1, arrM, arrM1, intMeth1, intMeth, 12, boolListAll)

                        'dv = dgv1.DataSource
                        'int2 = FindRowDV("Sample Size", dv)
                        'int3 = FindRowDV("Sample Size Units", dv)
                        'If intMeth = 1 Then
                        '    var2 = NZ(dv(int2).Item(arrM(2, 1)), "")
                        '    var3 = NZ(dv(int3).Item(arrM(2, 1)), ChrW(956) & "L")
                        '    varE = var2 & " " & var3
                        'Else
                        '    str1 = ""
                        '    For Count2 = 1 To intMeth
                        '        var2 = NZ(dv(int2).Item(arrM(2, Count2)), "")
                        '        var3 = NZ(dv(int3).Item(arrM(2, Count2)), ChrW(956) & "L")
                        '        str1 = str1 & var2 & " " & var3 & " for " & arrM(1, Count2) & Chr(10)
                        '    Next
                        '    'remove last line return
                        '    varE = Mid(str1, 1, Len(str1) - 1)
                        'End If

                        'Case "Analytical Method Type" '20190212 LEE: deprecated
                        '    varE = frmH.cbxAssayTechniqueAcronym.Text
                    Case "Assay Technique" '20190212 LEE:
                        varE = frmH.cbxAssayTechniqueAcronym.Text
                    Case "Data Archival"
                        dv = frmH.dgvDataCompany.DataSource
                        int1 = FindRowDV("Data Archival Location", dv)
                        varE = NZ(dv(int1).Item("Value"), "")
                    Case "Analytes"
                        strF = "IsIntStd = 'No'"
                        strS = "AnalyteDescription"
                        Erase rowsX
                        rowsX = tblAnalytesHome.Select(strF, strS)
                        int1 = rowsX.Length
                        str1 = ""
                        For Count2 = 0 To int1 - 1
                            var1 = rowsX(Count2).Item("AnalyteDescription")
                            If Count2 = int1 - 1 Then
                                str1 = str1 & var1
                            Else
                                str1 = str1 & var1 & ChrW(10)
                            End If
                        Next
                        varE = NZ(str1, "NA")

                    Case "Internal Standard"

                        strF = "IsIntStd = 'No'"
                        strS = "AnalyteDescription"
                        Erase rowsX
                        rowsX = tblAnalytesHome.Select(strF, strS)
                        int1 = rowsX.Length
                        str1 = ""
                        For Count2 = 0 To int1 - 1
                            var1 = rowsX(Count2).Item("AnalyteDescription")
                            var2 = rowsX(Count2).Item("IntStd")
                            If Count2 = int1 - 1 Then
                                str1 = str1 & var2 & " for " & var1
                            Else
                                str1 = str1 & var2 & " for " & var1 & ChrW(10)
                            End If
                        Next
                        varE = NZ(str1, "NA")

                    Case "Standard Curve Dynamic Range"
                        dgv = frmH.dgvWatsonAnalRef
                        int1 = dgv.Columns.Count
                        dv = frmH.dgvWatsonAnalRef.DataSource
                        int2 = FindRowDV("Is Internal Standard?", dv)
                        int3 = FindRowDV("LLOQ", dv)
                        int4 = FindRowDV("ULOQ", dv)
                        int5 = FindRowDV("LLOQ Units", dv)

                        int6 = FindRowDV("Alternate Calibr/QC Std Units", frmH.dgvStudyConfig.DataSource)
                        strUnits = NZ(frmH.dgvStudyConfig(1, int6).Value, "")
                        str2 = ""
                        If ctAnalytes = 1 Then
                            For Count2 = 1 To int1 - 1
                                str1 = NZ(dv(int2).Item(Count2), "")
                                If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                                    Exit For
                                Else
                                    var1 = SigFigOrDecString(CDec(NZ(dv(int3).Item(Count2), 0)), LSigFig, False)
                                    var2 = SigFigOrDecString(CDec(NZ(dv(int4).Item(Count2), 0)), LSigFig, False)
                                    var3 = dv(int5).Item(Count2)

                                    If Len(strUnits) = 0 Or StrComp(strUnits, "[None]", CompareMethod.Text) = 0 Then
                                    Else
                                        var3 = strUnits
                                    End If

                                    var4 = dgv.Columns.Item(Count2).HeaderText
                                    str2 = var1 & " " & var3 & " to " & var2 & " " & var3 & " for " & var4
                                End If
                            Next
                            varE = str2
                        Else
                            For Count2 = 1 To int1 - 1
                                str1 = NZ(dv(int2).Item(Count2), "")
                                If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                                    Exit For
                                Else
                                    var1 = SigFigOrDecString(CDec(NZ(dv(int3).Item(Count2), 0)), LSigFig, False)
                                    var2 = SigFigOrDecString(CDec(NZ(dv(int4).Item(Count2), 0)), LSigFig, False)
                                    var3 = dv(int5).Item(Count2)

                                    If Len(strUnits) = 0 Or StrComp(strUnits, "[None]", CompareMethod.Text) = 0 Then
                                    Else
                                        var3 = strUnits
                                    End If

                                    var4 = dgv.Columns.Item(Count2).HeaderText
                                    str2 = str2 & var1 & " " & var3 & " to " & var2 & " " & var3 & " for " & var4 & Chr(10)
                                End If
                            Next
                            'remove last line return
                            If int1 = 1 Then
                            Else
                                varE = Mid(str2, 1, Len(str2) - 1)
                            End If
                        End If

                    Case "Regression Type"
                        dv = frmH.dgvWatsonAnalRef.DataSource
                        int5 = FindRowDV("Is Internal Standard?", dv)
                        int2 = FindRowDV("Internal Standard", dv)
                        int3 = FindRowDV("Regression", dv)
                        int4 = FindRowDV("Weighting", dv)
                        dgv = frmH.dgvWatsonAnalRef
                        int1 = dgv.Columns.Count
                        'find internal standards
                        str2 = ""
                        If ctAnalytes = 1 Then
                            For Count2 = 1 To int1 - 1

                                var2 = dgv.Columns.Item(Count2).HeaderText
                                var3 = NZ(dv(int3).Item(Count2), "NA") 'Regression
                                var4 = NZ(dv(int4).Item(Count2), "NA") 'weighting
                                str1 = NZ(dv(int2).Item(Count2), "") & " for " & var2

                                str3 = NZ(dv(int5).Item(Count2), "")
                                If StrComp(str3, "Yes", CompareMethod.Text) = 0 Then
                                    Exit For
                                Else
                                    var1 = var3 & " (Weighted: " & var4 & ") for " & var2
                                    If Count2 = 1 Then
                                        str2 = var1
                                    Else
                                        str2 = str2 & ChrW(10) & var1
                                    End If

                                End If
                            Next
                            varE = str2
                        Else
                            For Count2 = 1 To int1 - 1
                                str1 = NZ(dv(int2).Item(Count2), "")
                                var4 = NZ(dv(int4).Item(Count2), "NA") 'weighting
                                str3 = NZ(dv(int5).Item(Count2), "")
                                If StrComp(str3, "Yes", CompareMethod.Text) = 0 Then
                                    Exit For
                                Else
                                    var1 = NZ(dv(int3).Item(Count2), "NA")
                                    var2 = dgv.Columns.Item(Count2).HeaderText
                                    str2 = str2 & var1 & " (Weighted: " & var4 & ") for " & var2 & Chr(10)
                                End If
                            Next
                            'remove last line return

                            If int1 = 1 Then
                            Else
                                varE = Mid(str2, 1, Len(str2) - 1)
                            End If
                        End If

                    Case "Total Number of Samples Analyzed"
                        If boolVal Then
                            varE = "NA"
                        Else
                            varE = frmH.txtSRecTotalReport.Text
                        End If

                    Case "Storage Conditions"

                        dv = frmH.dgvSampleReceipt.DataSource
                        int1 = dv.Count
                        If int1 = 0 Then
                            varE = ""
                        Else
                            var1 = NZ(dv(0).Item("CHARSTORAGETEMP"), "[None]")
                            'var2 = Capit(var1)
                            If Len(var1) = 0 Then
                                var2 = ""
                            Else
                                'check to se if should capitalized
                                var2 = Mid(var1, 1, 1)
                                If IsNumeric(var2) Then
                                    var2 = var1
                                Else
                                    var2 = Capit(var1)
                                End If
                            End If
                            varE = var2
                        End If

                    Case "First Samples Received"
                        If boolVal Then
                            varE = "NA"
                        Else
                            dv = frmH.dgvSampleReceipt.DataSource
                            int1 = dv.Count
                            If int1 = 0 Then
                                varE = ""
                            Else
                                'var1 = CDate(dv(0).Item("DTSHIPMENTRECEIVED"))
                                var2 = dv(0).Item("DTSHIPMENTRECEIVED")
                                'var1 = CDate(var2)
                                If Len(NZ(var2, "")) = 0 Then
                                    varE = "" 'System.DBNull.Value
                                Else
                                    var1 = CDate(var2)
                                    varE = Format(var1, LDateFormat)
                                End If
                            End If
                        End If

                    Case "Last Samples Analyzed"
                        If boolVal Then
                            varE = "NA"
                        Else
                            dv = frmH.dgvDataWatson.DataSource
                            int1 = FindRowDV("Last Extraction Date", dv)
                            var1 = NZ(dv(int1).Item("Value"), "[None]")
                            If Len(var1) = 0 Then
                                varE = ""
                            Else


                                Try
                                    varE = Format(CDate(var1), LDateFormat)
                                Catch ex As Exception
                                    varE = ""
                                End Try
                            End If
                        End If


                    Case "Demonstrated Freeze/thaw Stability Cycles"
                        dv = dgv1.DataSource
                        int2 = FindRowDV("Demonstrated Freeze/Thaw Cycles", dv)
                        If intMeth = 1 Then
                            var2 = NZ(dv(int2).Item(arrM(1, 1)), "[None]")
                            varE = var2
                        Else
                            str1 = ""
                            For Count2 = 1 To intMeth
                                var2 = NZ(dv(int2).Item(arrM(1, Count2)), "[None]")
                                str1 = str1 & var2 & " for " & arrM(1, Count2) & Chr(10)
                            Next
                            'remove last line return
                            varE = Mid(str1, 1, Len(str1) - 1)
                        End If

                    Case "Maximum # of Freeze/thaw Cycles"
                        dv = dgv1.DataSource
                        int2 = FindRowDV("Maximum # of Freeze/thaw Cycles", dv)
                        If intMeth = 1 Then
                            var2 = NZ(dv(int2).Item(arrM(1, 1)), "[None]")
                            varE = var2

                        Else
                            str1 = ""
                            For Count2 = 1 To intMeth
                                var2 = NZ(dv(int2).Item(arrM(2, Count2)), "[None]")
                                str1 = str1 & var2 & " for " & arrM(1, Count2) & Chr(10)
                            Next
                            'remove last line return
                            varE = Mid(str1, 1, Len(str1) - 1)
                        End If

                    Case "Stability under Storage Conditions" 'deprecated, is now BenchTopStability 20181110
                    Case "Bench-top Stability"
                        dv = dgv1.DataSource
                        ' int2 = FindRowDV("Stability Under Storage Conditions", dv)
                        int2 = FindRowDV("Bench-top Stability", dv)
                        If intMeth = 1 Then
                            var2 = NZ(dv(int2).Item(arrM(1, 1)), "[None]")
                            varE = var2
                        Else
                            str1 = ""
                            For Count2 = 1 To intMeth
                                var2 = NZ(dv(int2).Item(arrM(2, Count2)), "[None]")
                                str1 = str1 & var2 & " for " & arrM(1, Count2) & Chr(10)
                            Next
                            'remove last line return
                            varE = Mid(str1, 1, Len(str1) - 1)
                        End If

                    Case "Is Stability >= Maximum Storage Duration"
                        dv = dgv1.DataSource
                        int2 = FindRowDV("Is Stability >= Maximum Storage Duration", dv)
                        If intMeth = 1 Then
                            var2 = NZ(dv(int2).Item(arrM(1, 1)), "[None]")
                            varE = var2
                        Else
                            str1 = ""
                            For Count2 = 1 To intMeth
                                var2 = NZ(dv(int2).Item(arrM(2, Count2)), "[None]")
                                str1 = str1 & var2 & " for " & arrM(1, Count2) & Chr(10)
                            Next
                            'remove last line return
                            varE = Mid(str1, 1, Len(str1) - 1)
                        End If

                    Case "Maximum Run Size"
                        dv = dgv1.DataSource
                        int2 = FindRowDV("Maximum Run Size", dv)
                        var2 = NZ(dv(int2).Item(arrM(2, 1)), "[None]")
                        varE = var2

                    Case Else
                        dv = dgv1.DataSource
                        str1 = varR
                        varE = GetSummaryStringAll(dv, str1, arrM)


                End Select
            End If
            rows(Count1).Item("charValue") = varE
            rows(Count1).EndEdit()
        Next
        dgv = frmH.dgvSummaryData
        'find index of charvalue
        'int1 = dgv.Columns.Item("CHARVALUE").DisplayIndex
        Try
            int1 = dgv.Columns.Item("CHARVALUE").DisplayIndex
            dgv.AutoResizeColumn(int1)
            dgv.AutoResizeColumns()
        Catch ex As Exception

        End Try


    End Sub

    Function GetSummaryStringAll(ByVal dv As System.Data.DataView, ByVal strRow As String, ByVal arrM As Array) As String


        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2, var3, var4
        Dim varE
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim intHit As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        GetSummaryStringAll = ""
        int2 = FindRowDV(strRow, dv)
        str1 = ""

        str1 = ""
        Try
            For Count2 = 1 To ctAnalytes
                str2 = arrAnalytes(1, Count2)
                int1 = arrM(2, Count2)
                var2 = NZ(dv(int2).Item(int1), "[None]")
                If Len(var2) = 0 Then
                    var3 = var2
                Else
                    var3 = var2

                End If
                str1 = str1 & var3 & " for " & str2 & Chr(10)
            Next

            'remove last line return
            GetSummaryStringAll = Mid(str1, 1, Len(str1) - 1)
        Catch ex As Exception
            GetSummaryStringAll = "Error"
        End Try



    End Function


    Sub UpdateBoolSummaryTable()
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim int1 As Short
        Dim Count1 As Short
        Dim var1

        strF = "id_tblStudies = " & id_tblStudies
        tbl = tblSummaryData
        rows = tbl.Select(strF)
        int1 = rows.Length
        For Count1 = 0 To int1 - 1
            rows(Count1).BeginEdit()
            var1 = NZ(rows(Count1).Item("boolInclude"), "")
            If Len(var1) = 0 Then
                rows(Count1).Item("boolI") = True
            ElseIf var1 = 0 Then
                rows(Count1).Item("boolI") = False
            Else
                rows(Count1).Item("boolI") = True
            End If
            rows(Count1).EndEdit()
        Next

    End Sub

    Sub DoCancelSummaryTab()

        tblSummaryData.RejectChanges()

        Call ConfigureSummaryTable()


    End Sub


    Sub SaveSummaryData()

        frmH.dgvSummaryData.CommitEdit(DataGridViewDataErrorContexts.Commit)


        Dim dvCheck As System.Data.DataView = New DataView(tblSummaryData)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int1 As Short
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblSummaryData)

            If boolGuWuOracle Then

                Try
                    ta_tblSummaryData.Update(tblSummaryData)
                Catch ex As DBConcurrencyException
                    Try
                        'ds2005.TBLSUMMARYDATA.Merge('ds2005.TBLSUMMARYDATA, True)
                    Catch ex1 As Exception
                        MsgBox("TBLSUMMARYDATA: " & ex1.Message)

                    End Try

                End Try

            ElseIf boolGuWuAccess Then
                Try
                    ta_tblSummaryDataAcc.Update(tblSummaryData)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLSUMMARYDATA.Merge('ds2005Acc.TBLSUMMARYDATA, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblSummaryDataSQLServer.Update(tblSummaryData)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLSUMMARYDATA.Merge('ds2005Acc.TBLSUMMARYDATA, True)
                End Try
            End If

        End If

        Call UpdateBoolSummaryTable()

        frmH.dgvSummaryData.AutoResizeColumns()


    End Sub

    Sub AddAnalyteColReportTableAnalytes()

        If tblReportTableAnalytes.Columns.Contains("CHARANALYTE") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "CHARANALYTE"
            col1.DataType = System.Type.GetType("System.String")
            col1.Caption = "CHARANALYTE"
            col1.AllowDBNull = True
            tblReportTableAnalytes.Columns.Add(col1)
        End If

        'add analytes
        Dim var1, var2, var3
        Dim Count1 As Short
        Dim strF As String
        Dim strF1 As String
        Dim rows() As DataRow
        Dim Count2 As Short
        Dim str1 As String
        Dim int1 As Int16

        'Legend
        'Dim arrAnalytes(14, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        ''10=UseIntStd, 11=IntStd, 12=MasterAssayID,13=IsCoadministeredCmpd,14=Original Analyte Description

        'tblCalStdGroupAssayIDsAll
        If boolUseGroups Then

            Try
                For Count1 = 0 To tblAnalyteGroups.Rows.Count - 1
                    str1 = tblAnalyteGroups.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
                    int1 = tblAnalyteGroups.Rows(Count1).Item("INTGROUP")
                    strF1 = "INTGROUP = " & int1
                    'Dim rowsA() As DataRow = tblCalStdGroupsAcc.Select(strF1)
                    Dim rowsA() As DataRow = tblCalStdGroupAssayIDsAll.Select(strF1)

                    var1 = rowsA(0).Item("ANALYTEID")
                    var2 = rowsA(0).Item("ANALYTEINDEX")
                    var3 = rowsA(0).Item("MASTERASSAYID")

                    'strF = "ANALYTEID = " & var1 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND ID_TBLSTUDIES = " & id_tblStudies
                    strF = "INTGROUP = " & int1 & " AND ID_TBLSTUDIES = " & id_tblStudies
                    Dim rowsB() As DataRow = tblReportTableAnalytes.Select(strF)
                    For Count2 = 0 To rowsB.Length - 1
                        rowsB(Count2).BeginEdit()
                        rowsB(Count2).Item("CHARANALYTE") = str1
                        rowsB(Count2).EndEdit()
                    Next

                Next
            Catch ex As Exception
                var1 = var1 'debug
                var2 = ex.Message
                var2 = var2
            End Try

            var1 = var1 'debug

        Else

            For Count1 = 1 To ctAnalytes
                var1 = arrAnalytes(2, Count1) 'analyteid
                var2 = arrAnalytes(3, Count1) 'analyteindex
                var3 = arrAnalytes(12, Count1) 'masterassayid

                strF = "ANALYTEID = " & var1 & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND ID_TBLSTUDIES = " & id_tblStudies
                Erase rows
                rows = tblReportTableAnalytes.Select(strF)
                For Count2 = 0 To rows.Length - 1
                    rows(Count2).BeginEdit()
                    rows(Count2).Item("CHARANALYTE") = arrAnalytes(1, Count1)
                    rows(Count2).EndEdit()
                Next
            Next

        End If

        tblReportTableAnalytes.AcceptChanges()


    End Sub

    Function SaveTableReportData() As Boolean

        SaveTableReportData = False

        Dim Count1 As Short
        Dim Count2 As Short
        Dim dtbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim intCols As Short
        Dim intRows As Short
        Dim var1, var2, var3, var4, var5, var6, var7, var8
        Dim varID
        Dim tblAnal As System.Data.DataTable
        Dim tblRTables As System.Data.DataTable
        Dim str1 As String
        Dim dvAnalytes As System.Data.DataView
        Dim drows1() As DataRow
        Dim drows2() As DataRow
        Dim drowsMaxID() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim strF As String
        Dim maxID
        Dim maxID1
        'Dim arrMaxID(100) As Int64

        Dim intGroup As Short

        'find maxID for tblReportTable
        maxID = GetMaxID("tblReportTable", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid

        'str1 = "charTable = 'tblReportTable'"
        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If
        'drowsMaxID = tblMaxID.Select(str1)
        'maxID = drowsMaxID(0).Item("numMaxID")
        maxID1 = maxID

        dtbl = tblReportTables
        tblAnal = tblReportTableAnalytes
        tblRTables = tblReportTable
        intCols = dtbl.Columns.Count
        intRows = dtbl.Rows.Count

        Dim dgv As DataGridView = frmH.dgvReportTableConfiguration
        dv = dgv.DataSource

        'Try

        '    dv.RowStateFilter = DataViewRowState.Added
        '    If dv.Count > 0 Then
        '        var1 = dv(0).Item("CHARFCID")
        '    End If
        '    dv.RowStateFilter = DataViewRowState.CurrentRows
        '    If dv.Count > 0 Then
        '        var2 = dv(0).Item("CHARFCID")
        '    End If
        '    dv.RowStateFilter = DataViewRowState.Deleted
        '    If dv.Count > 0 Then
        '        var3 = dv(0).Item("CHARFCID")
        '    End If
        '    dv.RowStateFilter = DataViewRowState.ModifiedCurrent
        '    If dv.Count > 0 Then
        '        var4 = dv(0).Item("CHARFCID")
        '    End If
        '    dv.RowStateFilter = DataViewRowState.ModifiedOriginal
        '    If dv.Count > 0 Then
        '        var5 = dv(0).Item("CHARFCID")
        '    End If
        '    dv.RowStateFilter = DataViewRowState.None
        '    If dv.Count > 0 Then
        '        var6 = dv(0).Item("CHARFCID")
        '    End If
        '    dv.RowStateFilter = DataViewRowState.OriginalRows
        '    If dv.Count > 0 Then
        '        var7 = dv(0).Item("CHARFCID")
        '    End If
        '    dv.RowStateFilter = DataViewRowState.Unchanged
        '    If dv.Count > 0 Then
        '        var8 = dv(0).Item("CHARFCID")
        '    End If

        '    var8 = var8

        'Catch ex As Exception
        '    var8 = var8
        'End Try

        'Dim rowsOR() As DataRow = tblReportTable.Select("", "", DataViewRowState.CurrentRows)

        'Try
        '    Dim rows11() As DataRow = tblRTables.Select("", "", DataViewRowState.Added)
        '    Dim rows12() As DataRow = tblRTables.Select("", "", DataViewRowState.CurrentRows)
        '    Dim rows13() As DataRow = tblRTables.Select("", "", DataViewRowState.Deleted)
        '    Dim rows14() As DataRow = tblRTables.Select("", "", DataViewRowState.ModifiedCurrent)
        '    Dim rows15() As DataRow = tblRTables.Select("", "", DataViewRowState.ModifiedOriginal)
        '    Dim rows16() As DataRow = tblRTables.Select("", "", DataViewRowState.None)
        '    Dim rows17() As DataRow = tblRTables.Select("", "", DataViewRowState.OriginalRows)
        '    Dim rows18() As DataRow = tblRTables.Select("", "", DataViewRowState.Unchanged)

        '    If rows11.Length = 0 Then
        '    Else
        '        var1 = rows11(0).Item("CHARFCID")
        '    End If
        '    If rows12.Length = 0 Then
        '    Else
        '        var2 = rows12(0).Item("CHARFCID")
        '    End If
        '    If rows13.Length = 0 Then
        '    Else
        '        var3 = rows13(0).Item("CHARFCID")
        '    End If
        '    If rows14.Length = 0 Then
        '    Else
        '        var4 = rows14(0).Item("CHARFCID")
        '    End If
        '    If rows15.Length = 0 Then
        '    Else
        '        var5 = rows15(0).Item("CHARFCID")
        '    End If
        '    If rows16.Length = 0 Then
        '    Else
        '        var6 = rows16(0).Item("CHARFCID")
        '    End If
        '    If rows17.Length = 0 Then
        '    Else
        '        var7 = rows17(0).Item("CHARFCID")
        '    End If
        '    If rows18.Length = 0 Then
        '    Else
        '        var8 = rows18(0).Item("CHARFCID")
        '    End If

        '    var8 = var8
        'Catch ex As Exception
        '    var8 = var8
        'End Try


        Dim arrMaxID(intRows) As Int64
        '
        For Count1 = 0 To intRows - 1
            'find id_tblReportTable
            var1 = dtbl.Rows.Item(Count1).Item("id_tblConfigReportTables") 'id_tblConfigReportTables
            str1 = "id_tblStudies = " & id_tblStudies & " and id_tblConfigReportTables = " & var1

            var1 = dtbl.Rows.Item(Count1).Item("id_tblReportTable")
            str1 = "id_tblReportTable = " & var1
            drows1 = tblRTables.Select(str1)

            If drows1.Length = 0 Then 'add a new row
                Dim drow As DataRow = tblRTables.NewRow()
                maxID = dtbl.Rows.Item(Count1).Item("id_tblReportTable") 'maxID + 1
                drow.BeginEdit()
                drow.Item("id_tblReportTable") = maxID
                drow.Item("id_tblStudies") = id_tblStudies
                drow.Item("id_tblConfigReportTables") = dtbl.Rows.Item(Count1).Item("id_tblConfigReportTables")
                'drow.Item("boolRequiresSampleAssignment") = dtbl.Rows.item(Count1).Item("boolRequiresSampleAssignment")
                var1 = dtbl.Rows.Item(Count1).Item("boolRequiresSampleAssignment")
                If var1 Then
                    drow.Item("boolRequiresSampleAssignment") = -1 'dtbl.Rows.item(Count1).Item("Include")
                Else
                    drow.Item("boolRequiresSampleAssignment") = 0 'dtbl.Rows.item(Count1).Item("Include")
                End If
                var1 = dtbl.Rows.Item(Count1).Item("boolInclude")
                If var1 Then
                    drow.Item("boolInclude") = -1 'dtbl.Rows.item(Count1).Item("Include")
                Else
                    drow.Item("boolInclude") = 0 'dtbl.Rows.item(Count1).Item("Include")
                End If
                drow.Item("intOrder") = dtbl.Rows.Item(Count1).Item("intOrder")
                drow.Item("charPageOrientation") = dtbl.Rows.Item(Count1).Item("charPageOrientation")
                drow.Item("CHARSTABILITYPERIOD") = dtbl.Rows.Item(Count1).Item("CHARSTABILITYPERIOD")
                drow.Item("CHARSTYLE") = "Style 1"
                drow.Item("CHARHEADINGTEXT") = dtbl.Rows.Item(Count1).Item("CHARHEADINGTEXT")
                drow.Item("INTEGNUM") = dtbl.Rows.Item(Count1).Item("INTEGNUM")
                drow.Item("CHARFCID") = dtbl.Rows.Item(Count1).Item("CHARFCID")
                Try
                    var1 = NZ(dtbl.Rows.Item(Count1).Item("BOOLPLACEHOLDER"), 0)
                    If var1 Then
                        drow.Item("BOOLPLACEHOLDER") = -1 'dtbl.Rows.item(Count1).Item("Include")
                    Else
                        drow.Item("BOOLPLACEHOLDER") = 0 'dtbl.Rows.item(Count1).Item("Include")
                    End If
                Catch ex As Exception
                    var1 = var1 'debug
                End Try

                drow.EndEdit()
                tblRTables.Rows.Add(drow)
                arrMaxID(Count1) = maxID
            Else
                drows1(0).BeginEdit()
                drows1(0).Item("intOrder") = dtbl.Rows.Item(Count1).Item("intOrder")
                drows1(0).Item("charPageOrientation") = dtbl.Rows.Item(Count1).Item("charPageOrientation")
                var1 = dtbl.Rows.Item(Count1).Item("boolRequiresSampleAssignment")
                If var1 Then
                    drows1(0).Item("boolRequiresSampleAssignment") = -1 'dtbl.Rows.item(Count1).Item("Include")
                Else
                    drows1(0).Item("boolRequiresSampleAssignment") = 0 'dtbl.Rows.item(Count1).Item("Include")
                End If

                Try
                    var1 = NZ(dtbl.Rows.Item(Count1).Item("BOOLPLACEHOLDER"), 0)
                    If var1 Then
                        drows1(0).Item("BOOLPLACEHOLDER") = -1 'dtbl.Rows.item(Count1).Item("Include")
                    Else
                        drows1(0).Item("BOOLPLACEHOLDER") = 0 'dtbl.Rows.item(Count1).Item("Include")
                    End If
                Catch ex As Exception
                    var1 = var1 'debug
                End Try


                var1 = NZ(dtbl.Rows.Item(Count1).Item("boolInclude"), 0)
                If var1 Then
                    drows1(0).Item("boolInclude") = -1 'dtbl.Rows.item(Count1).Item("Include")
                Else
                    drows1(0).Item("boolInclude") = 0 'dtbl.Rows.item(Count1).Item("Include")
                End If
                var1 = NZ(dtbl.Rows.Item(Count1).Item("BOOLPLACEHOLDER"), 0)
                If var1 Then
                    drows1(0).Item("BOOLPLACEHOLDER") = -1 'dtbl.Rows.item(Count1).Item("Include")
                Else
                    drows1(0).Item("BOOLPLACEHOLDER") = 0 'dtbl.Rows.item(Count1).Item("Include")
                End If
                drows1(0).Item("CHARSTABILITYPERIOD") = dtbl.Rows.Item(Count1).Item("CHARSTABILITYPERIOD")
                'ENTER CODE HERE FOR:
                '  CHARHEADINGTEXT
                drows1(0).Item("CHARHEADINGTEXT") = dtbl.Rows.Item(Count1).Item("CHARHEADINGTEXT")
                '  CHARSTYLE
                drows1(0).Item("CHARSTYLE") = "Style 1"
                drows1(0).Item("INTEGNUM") = dtbl.Rows.Item(Count1).Item("INTEGNUM")
                drows1(0).Item("CHARFCID") = dtbl.Rows.Item(Count1).Item("CHARFCID")

                drows1(0).EndEdit()
                var1 = drows1(0).Item("id_tblReportTable")
                arrMaxID(Count1) = var1 'drows1(0).Item("id_tblReportTable")
            End If
        Next

        SaveTableReportData = True

        'Dim rowsO() As DataRow
        'If rowsOR.Length = 0 Then
        'Else
        '    'evaluate CHARFCID
        '    For Count1 = 0 To rowsOR.Length - 1
        '        var1 = rowsOR(Count1).Item("ID_TBLREPORTTABLE")
        '        var2 = NZ(rowsOR(Count1).Item("CHARFCID"), "")
        '        var3 = rowsOR(Count1).Item("BOOLINCLUDE")
        '        strF = "ID_TBLREPORTTABLE = " & var1
        '        rowsO = tblReportTable.Select(strF, "", DataViewRowState.ModifiedCurrent)
        '        If rowsO.Length = 0 Then
        '        Else
        '            var4 = NZ(rowsO(0).Item("CHARFCID"), "")
        '            var5 = rowsO(0).Item("BOOLINCLUDE")
        '            If StrComp(var2, var4, CompareMethod.Text) = 0 Then
        '                'check if boolinclude has changed
        '                If var3 = var5 Then
        '                Else
        '                    SaveTableReportData = True
        '                    Exit For
        '                End If
        '            Else
        '                SaveTableReportData = True
        '                Exit For
        '            End If
        '        End If
        '    Next

        'End If

        Call FillAuditTrailTemp(tblReportTable)

        Try
            If boolGuWuOracle Then
                Try
                    ta_tblReportTable.Update(tblReportTable)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLREPORTTABLE.Merge('ds2005.TBLREPORTTABLE, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportTableAcc.Update(tblReportTable)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTTABLE.Merge('ds2005Acc.TBLREPORTTABLE, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportTableSQLServer.Update(tblReportTable)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTTABLE.Merge('ds2005Acc.TBLREPORTTABLE, True)
                End Try
            End If
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


        tblRTables = tblReportTable

        'record new maxid
        If maxID = maxID1 Then
        Else

            Call PutMaxID("tblReportTable", maxID)

            'drowsMaxID(0).BeginEdit()
            'drowsMaxID(0).Item("numMaxID") = maxID
            'drowsMaxID(0).EndEdit()
            'Try
            '    If boolGuWuOracle Then
            '        Try
            '            ta_tblMaxID.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuAccess Then
            '        Try
            '            ta_tblMaxIDAcc.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    ElseIf boolGuWuSQLServer Then
            '        Try
            '            ta_tblMaxIDSQLServer.Update(tblMaxID)
            '        Catch ex As DBConcurrencyException
            '            'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
            '        End Try
            '    End If

            'Catch ex As Exception
            '    var1 = ex.Message
            '    var1 = var1
            'End Try

        End If


        'now fill Analyte checkboxes
        maxID = GetMaxID("TBLREPORTTABLEANALYTES", 1, False)
        Dim maxIDO As Int64
        maxIDO = maxID
        For Count1 = 0 To intRows - 1
            'find id_tblReportTable
            var1 = dtbl.Rows.Item(Count1).Item("id_tblConfigReportTables") 'id_tblConfigReportTables
            str1 = "id_tblStudies = " & id_tblStudies & " and id_tblConfigReportTables = " & var1
            drows1 = tblRTables.Select(str1)
            varID = arrMaxID(Count1)

            For Count2 = 1 To ctAnalytes
                var1 = arrAnalytes(2, Count2) 'AnalyteID
                var2 = arrAnalytes(3, Count2) 'AnalyteIndex
                var3 = arrAnalytes(12, Count2) 'MasterAssayID

                If boolUseGroups Then
                    intGroup = NZ(arrAnalytes(15, Count2), 0) 'group
                    strF = "id_tblStudies = " & id_tblStudies & " AND INTGROUP = " & intGroup & " AND id_tblReportTable = " & varID
                Else
                    intGroup = 0
                    strF = "id_tblStudies = " & id_tblStudies & " AND ANALYTEID = " & var1 & " AND id_tblReportTable = " & varID & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3
                End If
                drows2 = tblAnal.Select(strF) 'tblReportTableAnalytes
                Try
                    If drows2.Length = 0 Then 'add a new row
                        maxID = maxID + 1
                        Dim drowA As DataRow = tblAnal.NewRow()
                        drowA.Item("ID_TBLREPORTTABLEANALYTES") = maxID
                        drowA.Item("id_tblStudies") = id_tblStudies
                        drowA.Item("ANALYTEID") = var1
                        drowA.Item("ANALYTEINDEX") = var2
                        drowA.Item("MASTERASSAYID") = var3
                        drowA.Item("id_tblReportTable") = varID
                        var3 = arrAnalytes(1, Count2)
                        If dtbl.Columns.Contains(var3) Then
                            var2 = dtbl.Rows.Item(Count1).Item(var3) 'tblReportTables
                            If var2 Then
                                drowA.Item("boolInclude") = -1
                            Else
                                drowA.Item("boolInclude") = 0
                            End If
                        End If

                        drowA.Item("INTGROUP") = intGroup
                        tblAnal.Rows.Add(drowA)
                    Else '
                        var3 = arrAnalytes(1, Count2)
                        If dtbl.Columns.Contains(var3) Then
                            var2 = dtbl.Rows.Item(Count1).Item(var3) 'tblReportTables
                            If var2 Then
                                drows2(0).Item("boolInclude") = -1
                            Else
                                drows2(0).Item("boolInclude") = 0
                            End If
                        End If

                        drows2(0).Item("INTGROUP") = intGroup
                        drows2(0).EndEdit()
                    End If
                Catch ex As Exception
                    Dim ve
                    ve = ex.Message
                End Try

            Next
        Next

        If maxIDO = maxID Then 'ignore
        Else
            Call PutMaxID("TBLREPORTTABLEANALYTES", maxID)
        End If

        Dim dvCheck As System.Data.DataView = New DataView(tblReportTableAnalytes)
        dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        Dim int10 As Short
        int10 = 1
        If int10 = 0 Then
        Else

            Call FillAuditTrailTemp(tblReportTableAnalytes)

            Try
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
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try



        End If


    End Function

    Sub FillTableReportDataAnalytes(ByVal boolInitial As Boolean)

        Dim Count1 As Short
        Dim Count2 As Short
        Dim dtbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim intCols As Short
        Dim intRows As Short
        Dim var1, var2, var3, var4, var5, var6
        Dim varID
        Dim tblAnalytes As System.Data.DataTable
        Dim tblRTables As System.Data.DataTable
        Dim drowsR() As DataRow
        Dim str1 As String
        Dim dvAnalytes As System.Data.DataView
        Dim drows1() As DataRow
        Dim drows2() As DataRow
        Dim int1 As Short
        Dim int2 As Short
        Dim strF As String
        Dim strO As String
        Dim bool As Short
        Dim tblC As System.Data.DataTable
        Dim rowsC() As DataRow
        Dim idC As Long
        Dim intGroup As Short

        'Dim cs As DataGridColumnStyle

        dtbl = tblReportTables
        tblC = tblConfigReportTables
        tblAnalytes = tblReportTableAnalytes
        tblRTables = tblReportTable
        intRows = dtbl.Rows.Count

        strF = "ID_TBLSTUDIES = " & id_tblStudies
        strO = "INTORDER ASC, CHARHEADINGTEXT ASC"
        drowsR = tblRTables.Select(strF, strO)

        Dim tbl As System.Data.DataTable
        Dim rowsT() As DataRow
        tbl = tblConfigReportTables

        intRows = dtbl.Rows.Count

        Dim boolF As Boolean
        For Count1 = 0 To intRows - 1

            'now fill Analyte checkboxes
            dtbl.Rows(Count1).BeginEdit()
            For Count2 = 1 To ctAnalytes
                var1 = arrAnalytes(2, Count2) 'AnalyteID
                var2 = arrAnalytes(3, Count2) 'AnalyteIndex
                var3 = arrAnalytes(12, Count2) 'MasterAssayID
                intGroup = NZ(arrAnalytes(15, Count2), 0) 'intGroup
                varID = dtbl.Rows(Count1).Item("id_tblReportTable")
                strF = "id_tblStudies = " & id_tblStudies & " AND ANALYTEID = " & var1 & " AND id_tblReportTable = " & varID & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3
                If boolUseGroups Then
                    strF = "id_tblStudies = " & id_tblStudies & " AND INTGROUP = " & intGroup & " AND id_tblReportTable = " & varID
                Else
                    strF = "id_tblStudies = " & id_tblStudies & " AND ANALYTEID = " & var1 & " AND id_tblReportTable = " & varID & " AND ANALYTEINDEX = " & var2 & " AND MASTERASSAYID = " & var3 & " AND INTGROUP = 0"
                End If
                drows2 = tblAnalytes.Select(strF)

                var1 = arrAnalytes(1, Count2)
                If drows2.Length = 0 Then 'configure as selected
                    'dtbl.Rows.item(Count1).Item(Count2 + 4) = True
                    'dtbl.Rows.Item(Count1).Item(arrAnalytes(1, Count2)) = True

                    If dtbl.Columns.Contains(var1) Then
                        dtbl.Rows(Count1).Item(var1) = True
                    End If

                Else
                    'dtbl.Rows.item(Count1).Item(Count2 + 4) = drows2(0).Item("boolInclude")
                    'dtbl.Rows.Item(Count1).Item(arrAnalytes(1, Count2)) = drows2(0).Item("boolInclude")
                    var2 = NZ(drows2(0).Item("boolInclude"), True)
                    If dtbl.Columns.Contains(var1) Then
                        dtbl.Rows(Count1).Item(var1) = var2
                    End If
                End If
            Next

            dtbl.Rows(Count1).EndEdit()

        Next

        'now ensure that any newly added tables in tblConfigReportTables are included here
        Dim tblMax As System.Data.DataTable
        Dim strFMax As String
        Dim rowsMax() As DataRow
        Dim maxID, maxID1

        'tblMax = tblMaxID

        'strFMax = "charTable = 'tblReportTable'"
        'Erase rowsMax
        'rowsMax = tblMax.Select(strFMax)
        'maxID = rowsMax(0).Item("nummaxid")

        maxID = GetMaxID("tblReportTable", 1, False) 'if maxid increment is 1, then getmaxid already does putmaxid
        maxID1 = maxID

        intRows = tblC.Rows.Count
        For Count1 = 0 To intRows - 1
            idC = tblC.Rows(Count1).Item("id_tblConfigReportTables")
            str1 = "id_tblConfigReportTables = " & idC
            Erase drows1
            drows1 = dtbl.Select(str1)
            If drows1.Length = 0 And idC < 1000 Then 'add a row

                maxID = maxID + 1
                Dim workRow As DataRow = dtbl.NewRow()
                workRow.BeginEdit()

                var1 = tblC.Rows(Count1).Item("boolRequiresSampleAssignment")
                workRow.Item("boolRequiresSampleAssignment") = var1

                idC = tblC.Rows(Count1).Item("id_tblConfigReportTables")
                workRow.Item("id_tblConfigReportTables") = idC

                var1 = tblC.Rows(Count1).Item("CHARTABLENAME")
                workRow.Item("CHARTABLENAME") = var1
                workRow.Item("CHARHEADINGTEXT") = var1

                var1 = False 'tblC.Rows(Count1).Item("BOOLINCLUDE")
                workRow.Item("BOOLINCLUDE") = var1

                var1 = tblC.Rows(Count1).Item("INTORDER")
                workRow.Item("INTORDER") = var1

                var1 = "P"
                workRow.Item("CHARPAGEORIENTATION") = var1

                var1 = System.DBNull.Value
                workRow.Item("CHARSTABILITYPERIOD") = var1

                'var1 = False 'tblC.Rows(Count1).Item("boolRequiresSampleAssignment")
                'workRow.Item("boolFigure") = var1

                var1 = "Style 1"
                workRow.Item("CHARSTYLE") = var1

                var1 = maxID
                workRow.Item("ID_TBLREPORTTABLE") = var1

                Try
                    var1 = tblC.Rows(Count1).Item("BOOLPLACEHOLDER")
                    workRow.Item("BOOLPLACEHOLDER") = var1
                Catch ex As Exception
                    var1 = var1
                End Try


                'now fill Analyte checkboxes
                For Count2 = 1 To ctAnalytes
                    var1 = arrAnalytes(1, Count2)
                    If dtbl.Columns.Contains(var1) Then
                        workRow.Item(arrAnalytes(1, Count2)) = True
                    End If
                Next

                workRow.EndEdit()
                dtbl.Rows.Add(workRow)

            End If
        Next


        If maxID = maxID1 Then 'ignore
        Else

            Call PutMaxID("tblReportTable", maxID)

            'rowsMax(0).BeginEdit()
            'rowsMax(0).Item("nummaxid") = maxID
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

        'set cmdOrderReportTableConfig Position
        Call SizecmdOrder(frmH.dgvReportTableConfiguration, frmH.cmdOrderReportTableConfig, "INTORDER")

        'frmH.cmdResize.Left = frmH.cmdOrderReportTableConfig.Left + frmH.cmdOrderReportTableConfig.Width + 10


        'If boolInitial Then
        '    dv = dtbl.DefaultView
        '    dv.Sort = "INTORDER ASC"
        '    If frmH.rbShowIncludedRTConfig.Checked Then
        '        strF = "Include = " & True 'Leave this as true
        '        dv.RowFilter = strF
        '    End If
        '    dv.AllowEdit = True
        '    dv.AllowNew = False
        '    dv.AllowDelete = False
        '    'frmh.dgReportTableConfiguration.DataSource = dv
        '    frmH.dgvReportTableConfiguration.DataSource = dv

        'End If
    End Sub


    Sub FillTableReportsAnalytes(ByVal boolInitial As Boolean)

        Dim Count1 As Short
        Dim Count2 As Short
        Dim dg As DataGrid
        Dim dtbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim boolRO As Boolean
        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2
        Dim tblAnalytes As System.Data.DataTable
        Dim str1 As String
        Dim dvAnalytes As System.Data.DataView
        Dim drows() As DataRow
        Dim intCols As Short
        Dim dgv As DataGridView
        Dim dtbl1 As System.Data.DataTable

        'Dim cs As DataGridColumnStyle
        'Note: a tablestyle already exists

        dgv = frmH.dgvReportTableConfiguration
        dtbl = tblReportTables
        tblAnalytes = tblReportTableAnalytes
        'boolRO = True
        boolRO = boolInitial
        'delete all columns except cols 0,1,2,3,4
        'delete all gridstyles except gs 0,1,2,3
        int1 = dtbl.Columns.Count
        Dim ts1 As DataGridTableStyle
        Dim cs As GridColumnStylesCollection
        'ts1 = dg.TableStyles(0)
        'cs = ts1.GridColumnStyles
        boolLoad = True 'to cancel a dgvReportTableConfig currentcell event
        dtbl1 = tblReportTable
        For Count2 = int1 - 1 To 0 Step -1
            str1 = dtbl.Columns.Item(Count2).ColumnName

            If dtbl.Columns.Contains(str1) Then 'ignore
            ElseIf StrComp(str1, "CHARTABLENAME", CompareMethod.Text) = 0 Then 'ignore
                dtbl.Columns.Remove(dtbl.Columns.Item(Count2))
            End If

            'If StrComp(str1, "boolRequiresSampleAssignment", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "CHARTABLENAME", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "INTORDER", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "CHARPAGEORIENTATION", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "id_tblConfigReportTables", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'ignore
            '    'ElseIf StrComp(str1, "boolFigure", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "CHARSTABILITYPERIOD", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "CHARHEADINGTEXT", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "CHARSTYLE", CompareMethod.Text) = 0 Then 'ignore
            'ElseIf StrComp(str1, "ID_TBLREPORTTABLE", CompareMethod.Text) = 0 Then 'ignore
            'Else
            '    dtbl.Columns.Remove(dtbl.Columns.Item(Count2))
            'End If
        Next
        intCols = dgv.Columns.Count

        'dgv.AutoGenerateColumns = True
        For Count2 = 1 To ctAnalytes
            Dim dc As New DataColumn
            dc.DataType = System.Type.GetType("System.Boolean")
            var1 = arrAnalytes(1, Count2)
            dc.ColumnName = arrAnalytes(1, Count2)
            dc.Caption = arrAnalytes(1, Count2)
            'dc.ReadOnly = boolRO
            dc.AllowDBNull = False
            dc.DefaultValue = False

            dtbl.Columns.Add(dc)
            'dgv.AutoResizeColumn(Count2 + intCols - 1)
            'dgv.Columns.Item(Count2 + intCols - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

        Next
        'dgv.AutoGenerateColumns = False

        Call SetComboCell(frmH.dgvReportTableConfiguration, "CHARPAGEORIENTATION")

        boolLoad = False

        'dv = dtbl.DefaultView
        Dim dv10 As System.Data.DataView = New DataView(dtbl)
        dv10.AllowNew = False
        dv10.AllowDelete = False
        'dg.DataSource = dv
        dgv.DataSource = dv10

        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()

        'dgv.Sort(dgv.Columns.Item("INTORDER"), System.ComponentModel.ListSortDirection.Ascending)


    End Sub

End Module
