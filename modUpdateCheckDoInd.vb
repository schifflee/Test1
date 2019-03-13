Option Compare Text

Module modUpdateCheckDoInd
    Public boolUpdateCheckBad As Boolean
    Public strUpdateMsg As String
    Public boolBad As Boolean = True
    Public boolNeedsUD As Boolean = True
    Public boolUpdateDB As Boolean = False
    Public boolQuitUpdate As Boolean = False

    Public conA As New ADODB.Connection
    Public conB As New ADODB.Connection
    Public cmd As New ADODB.Command
    Public rsVer As New ADODB.Recordset
    Public CurVer As String
    Public CurVerNew As String
    Public CurVerDB As String
    Public CurVerDB_New As String
    Public ver(4, 3) As Short
    Public intCount As Short
    Public intNV As Int32 = 0
    Public strNV_DB As String
    Public strNV_DB_New As String
    Public strNV_App As String


    Sub UpdatePG(ByRef intPG As Short, ByRef intPGMax As Short, ByRef frm As frmUpdateCheck, ByRef pg As ProgressBar)

        '20160801 LEE: This isn't working because pg is in a different thread
        'have to implement a BackGroundWorker, but too much work currently.
        'will just hide pg for now on frmUpdate

        intPG = intPG + 1
        If intPG >= intPGMax Then
            intPGMax = intPGMax + 10
            pg.Maximum = intPGMax
            pg.Refresh()
        End If
        pg.Value = intPG
        pg.Refresh()
        Application.DoEvents()


    End Sub

    Sub DoIndUpdates(ByRef frmUpdate As frmUpdateCheck, ByVal boolAccess As Boolean, ByVal boolSQLSer As Boolean, ByVal boolOra As Boolean, ByVal con As ADODB.Connection, ByVal cmd As OleDb.OleDbCommand, ByVal constr As String, ByRef pg As ProgressBar)
        'DoIndUpdates: Update StudyDoc Database (based on current StudyDoc Application Version Number vs. the previous StudyDoc Database Version Number)

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1
        Dim strNV As String

        Dim intPG As Short = 0
        Dim intPGMax As Short = 75

        'Try
        '    frmUpdate.pgOverall.Maximum = intPGMax
        '    frmUpdate.pgOverall.Value = 0
        '    frmUpdate.pgOverall.Refresh()
        'Catch ex As Exception
        '    var1 = ex.Message

        'End Try

        Try
            pg.Step = 1
            pg.Maximum = intPGMax
            pg.Value = 0
            pg.Refresh()
        Catch ex As Exception
            var1 = ex.Message

        End Try


        Cursor.Current = Cursors.WaitCursor

        str1 = frmUpdate.lbl1.Text


        intCount = 0

        '2.0.36
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 36)
        If strNV > strNV_DB Then
            intCount = intCount + 1
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_boolTheoretical"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_boolTheoretical(boolAccess, boolSQLSer, boolOra, con, cmd)
        End If


        '2.0.38
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 38)
        If strNV > strNV_DB Then
            intCount = intCount + 1
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_boolInclAnova"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_boolInclAnova(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblConfigReportTables_3132"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblConfigReportTables_3132(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblAssignedSamplesHelper_3132"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblAssignedSamplesHelper_3132(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.40
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 40)
        If strNV > strNV_DB Then
            'Update TBLASSIGNEDSAMPLESHELPER
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_INTEGNUM"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_INTEGNUM(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_BOOLBQLLEGEND"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_BOOLBQLLEGEND(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.41
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 41)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_CHARUNITS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_CHARUNITS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_CHARUNITS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_CHARUNITS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTAB_QCCALIBR"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTAB_QCCALIBR(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_MATRIX"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLCONFIGHEADERLOOKUP_MATRIX(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_SORT"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLTABLEPROPERTIES_SORT(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.44
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 44)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_tblTABLEPROPERTIES_BOOLINCLUDEPSAE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_tblTABLEPROPERTIES_BOOLINCLUDEPSAE(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.52
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 52)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTABLEPROPERTIES_BOOLCSREPORTACCVALUES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTABLEPROPERTIES_BOOLCSREPORTACCVALUES(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.58
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 58)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTAB1_REPORTTEMPLATE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTAB1_REPORTTEMPLATE(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.61
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 61)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_BIASDIFF"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLFIELDCODES_BIASDIFF(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.62
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 62)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_INTQCPERCDECPLACES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_INTQCPERCDECPLACES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_INTQCPERCDECPLACES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_INTQCPERCDECPLACES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLCONFIGURATION_INTQCPERCDECPLACES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLCONFIGURATION_INTQCPERCDECPLACES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLCONFIGURATION_DEFAULT1"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLCONFIGURATION_DEFAULT1(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_BOOLRC"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLTABLEPROPERTIES_BOOLRC(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFileds_TBLTABLEPROPERTIES_LEGENDS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFileds_TBLTABLEPROPERTIES_LEGENDS(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        'Don't do the Delete thing yet
        'Call DeleteRecords_TBLREPORTHEADERS( boolAccess, boolSQLSer, boolOra, con, cmd)

        '2.0.63
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 63)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTABLEPROPERTIES_CorrectBOOLPOSLEG"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTABLEPROPERTIES_CorrectBOOLPOSLEG(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_BOOLINCLUDEDATE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLTABLEPROPERTIES_BOOLINCLUDEDATE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_RECOVERY"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLTABLEPROPERTIES_RECOVERY(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.69
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 69)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_ANTICOAGULANTMETHOD"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLFIELDCODES_ANTICOAGULANTMETHOD(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_QAEVENTSBORDER"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_QAEVENTSBORDER(boolAccess, boolSQLSer, boolOra, con, cmd)
            'don't do this yet
            'Call AddFields_TBLQATABLES_CHARCRITICALPHASE( boolAccess, boolSQLSer, boolOra, con, cmd)
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLANALYTICALRUNSUMMARY_BOOLINCLUDEREGR"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLANALYTICALRUNSUMMARY_BOOLINCLUDEREGR(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.70
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 70)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_BOOLINCLUDEWATSONLABELS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLTABLEPROPERTIES_BOOLINCLUDEWATSONLABELS(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.0.73
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 0, 73)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLMETHODVALIDATIONDATA_CHARFTSTORCOND"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLMETHODVALIDATIONDATA_CHARFTSTORCOND(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_CHARFTSTORCOND_CHANGEORDER"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_CHARFTSTORCOND_CHANGEORDER(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_CHARFTSTORCOND"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_CHARFTSTORCOND(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_FREEZETHAWSTORAGECOND"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLFIELDCODES_FREEZETHAWSTORAGECOND(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 2)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLASSIGNEDSAMPLES_BOOLEXCLSAMPLE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLASSIGNEDSAMPLES_BOOLEXCLSAMPLE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLASSIGNEDSAMPLES_BOOLACCCRIT"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLASSIGNEDSAMPLES_BOOLACCCRIT(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLCONFIGURATION_EXCLUDESAMPLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLCONFIGURATION_EXCLUDESAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_EXCLUDESAMPLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_EXCLUDESAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_EXCLUDESAMPLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_EXCLUDESAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.5
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 5)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_boolInclAnovaSummaryStats"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_boolInclAnovaSummaryStats(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.7
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 7)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLANALREFSTANDARDS_CHARCHEMSTRUCTURE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLANALREFSTANDARDS_CHARCHEMSTRUCTURE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_CHARCHEMSTRUCTURE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_CHARCHEMSTRUCTURE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_CHARCHEMSTRUCTURE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_CHARCHEMSTRUCTURE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "DeleteRecords_TBLREPORTSTATEMENTS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call DeleteRecords_TBLREPORTSTATEMENTS(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.8
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 8)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_STUDYSTARTDATE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_STUDYSTARTDATE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_STUDYSTARTDATE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_STUDYSTARTDATE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_STUDYSTARTDATE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_STUDYSTARTDATE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTABLEPROPERTIES_BOOLSTATSRE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLTABLEPROPERTIES_BOOLSTATSRE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLCONFIGAPPFIGS_MISCELLANEOUS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLCONFIGAPPFIGS_MISCELLANEOUS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLAPPFIGS_CHARFCID"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLAPPFIGS_CHARFCID(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_TIMEPERIOD"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLTABLEPROPERTIES_TIMEPERIOD(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "DeleteRecords_TBLFIELDCODES_ANALYTE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call DeleteRecords_TBLFIELDCODES_ANALYTE(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.9
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 9)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblConfigReportTables_333435"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblConfigReportTables_333435(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblAssignedSamplesHelper_34"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblAssignedSamplesHelper_34(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblAssignedSamplesHelper_35"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblAssignedSamplesHelper_35(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.10
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 10)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTABLEPROPERTIES_MISCELLANEOUS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTABLEPROPERTIES_MISCELLANEOUS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLOUTSTANDINGITEMS_CHARFIELDCODE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLOUTSTANDINGITEMS_CHARFIELDCODE(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.12
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 12)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTAB1_REPORTTEMPLATE_A"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTAB1_REPORTTEMPLATE_A(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_CORRECTION"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_CORRECTION(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLREASONFORCHANGE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call CreateTable_TBLREASONFORCHANGE(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLMEANINGOFSIG"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call CreateTable_TBLMEANINGOFSIG(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLSAVEEVENT"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call CreateTable_TBLSAVEEVENT(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLAUDITTRAIL"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call CreateTable_TBLAUDITTRAIL(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLDATASYSTEM"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call CreateTable_TBLDATASYSTEM(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLSTUDIES_ID_TBLDATASYSTEM"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLSTUDIES_ID_TBLDATASYSTEM(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_COMPLIANCE_01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_COMPLIANCE_01(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTAB1_AUDITTRAIL"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLTAB1_AUDITTRAIL(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLMAXID_TBLAUDITTRAIL"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLMAXID_TBLAUDITTRAIL(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLFIELDCODES_NUMBEROFSAMPLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLFIELDCODES_NUMBEROFSAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.18
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 18)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_FORMATTHOUSAND"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_FORMATTHOUSAND(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_FORMATTHOUSAND"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_FORMATTHOUSAND(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblAssignedSamplesHelper_DILN"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblAssignedSamplesHelper_DILN(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.1.20
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 1, 20)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTAB1_COMPLIANCE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLTAB1_COMPLIANCE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLPERMISSIONS_COMPLIANCE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLPERMISSIONS_COMPLIANCE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLMAXID_AUDITTRAIL"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLMAXID_AUDITTRAIL(boolAccess, boolSQLSer, boolOra, con, cmd)
            'for some reason, previous code is throwing a big error

            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLCONFIGCOMPLIANCE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call CreateTable_TBLCONFIGCOMPLIANCE(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_PERMISSIONS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_PERMISSIONS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_COMPLIANCE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_COMPLIANCE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_CORRECTION02"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_CORRECTION02(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "DeleteRecords_TBLPERMISSIONS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call DeleteRecords_TBLPERMISSIONS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLPERMISSIONS_STUDYDESIGNNULLS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLPERMISSIONS_STUDYDESIGNNULLS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_RFC_MOS_DEFAULT"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_RFC_MOS_DEFAULT(frmUpdate, boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_CORPORATENICKNAMES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_CORPORATENICKNAMES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLSAVEEVENT_BOOLS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLSAVEEVENT_BOOLS(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.3
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 3)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_SIGBLOCKS01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLFIELDCODES_SIGBLOCKS01(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLREPORTS_BOOLDISPLAYATTACHMENT"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLREPORTS_BOOLDISPLAYATTACHMENT(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.4
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 4)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLREPORTS_BOOLINSERTWORDDOCS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLREPORTS_BOOLINSERTWORDDOCS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_SIGBLOCKS01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLFIELDCODES_TOCs(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.6
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 6)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLREPORTS_BOOLREADONLYTABLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLREPORTS_BOOLREADONLYTABLES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLAPPFIGS_BOOLINSERTWORDDOCS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLAPPFIGS_BOOLINSERTWORDDOCS(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.7
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 7)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblConfigReportTables_363738"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblConfigReportTables_363738(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_363738"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLCONFIGHEADERLOOKUP_363738(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_tblAssignedSamplesHelper_37"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_tblAssignedSamplesHelper_37(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_BOOLINTRARUNSTATS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLTABLEPROPERTIES_BOOLINTRARUNSTATS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_COMPLIANCE_02"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_COMPLIANCE_02(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_BOOLINCLANOVASUMSTATS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_BOOLINCLANOVASUMSTATS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTABLEPROPERTIES_BOOLINCLANOVASUMSTATS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTABLEPROPERTIES_BOOLINCLANOVASUMSTATS(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If


        '2.2.8
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 8)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_REDBOLDFONT"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_BOOLREDBOLDFONT(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_REDBOLDFONT"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_BOOLREDBOLDFONT(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.10
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 10)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTAB1_BOOLINCLUDEINTEMPLATE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTAB1_BOOLINCLUDEINTEMPLATE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTEMPLATEATTRIBUTES_13"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLTEMPLATEATTRIBUTES_13(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.14
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 14)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLREPORTTABLE_343536"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLREPORTTABLE_333435(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_ANOMALIES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_ANOMALIES(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.15
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 15)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AlterColumn_TBLMETHODVALIDATIONDATA_CHARPROCSTABILITY"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AlterColumn_TBLMETHODVALIDATIONDATA_CHARPROCSTABILITY(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.17
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 17)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_DTSTUDYSTARTDATE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_DTSTUDYSTARTDATE(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_NUMSIGFIGSAREA"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_NUMSIGFIGSAREA(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_NUMSIGFIGSAREA"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_NUMSIGFIGSAREA(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_ORDERNUMREGRDEC"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLDATATABLEROWTITLES_ORDERNUMREGRDEC(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLTABLEPROPERTIES_BOOLSTATSBIAS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLTABLEPROPERTIES_BOOLSTATSBIAS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AlterColumn_TBLQATABLES_DTCOLUMN1"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AlterColumn_TBLQATABLES_DTCOLUMN1(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.18
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 18)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_TOTALANALYTICALRUNS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLFIELDCODES_TOTALANALYTICALRUNS(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTAB1_CUSTOMFIELDCODES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLTAB1_CUSTOMFIELDCODES(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.19
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 19)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLCUSTOMFIELDCODE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call CreateTable_TBLCUSTOMFIELDCODE(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLMAXID_TBLCUSTOMFIELDCODES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLMAXID_TBLCUSTOMFIELDCODES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLFIELDCODES_BOOLCUSTOM"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLFIELDCODES_BOOLCUSTOM(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLPERMISSIONS_CUSTOMFIELDCODES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLPERMISSIONS_CUSTOMFIELDCODES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTEMPLATEATTRIBUTES_36"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLTEMPLATEATTRIBUTES_36(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLPERMISSIONS_NULL"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLPERMISSIONS_NULL(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_CUSTOMFIELDCODES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_CUSTOMFIELDCODES(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.22
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 22)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyAddRecords_TBLDATEFORMATS_01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyAddRecords_TBLDATEFORMATS_01(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLNOMCONCPAREN"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_BOOLNOMCONCPAREN(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLNOMCONCPAREN"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_BOOLNOMCONCPAREN(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.24
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 24)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLTABLEDTTIMESTAMP"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_BOOLTABLEDTTIMESTAMP(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLTABLEDTTIMESTAMP"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_BOOLTABLEDTTIMESTAMP(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.25
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 25)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLFOOTNOTEQCMEAN"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLDATATABLEROWTITLES_BOOLFOOTNOTEQCMEAN(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLFOOTNOTEQCMEAN"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLDATA_BOOLFOOTNOTEQCMEAN(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLREPORTTABLE_INCURREDSAMPLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call ModifyRecords_TBLREPORTTABLE_INCURREDSAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLREPORTTABLEANALYTES_NUMINCSAMPLECRIT01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddFields_TBLREPORTTABLEANALYTES_NUMINCSAMPLECRIT01(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_INCURREDSAMPLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLCONFIGHEADERLOOKUP_INCURREDSAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLREPORTTABLEHEADERCONFIG_INCURREDSAMPLES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Call AddRecords_TBLREPORTTABLEHEADERCONFIG_INCURREDSAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)

        End If

        '2.2.33
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 33)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLQCNA"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLDATATABLEROWTITLES_BOOLQCNA(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLQCNA"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLDATA_BOOLQCNA(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If


        '2.2.38
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 38)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_ANALYTE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLFIELDCODES_ANALYTE(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If

        '

        '2.2.40
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 40)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLBQL"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLDATA_BOOLBQL(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLBQL"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLDATATABLEROWTITLES_BOOLBQL(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
            'AddRecords_TBLDATATABLEROWTITLES_BOOLBQL
        End If

        '2.2.46
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 46)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "TBLREPORTTABLEHEADERCONFIG"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call Fix_TBLREPORTTABLEHEADERCONFIG(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If



        '2.2.51
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 51)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLIGNOREFC"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLDATA_BOOLIGNOREFC(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLIGNOREFC"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLDATATABLEROWTITLES_BOOLIGNOREFC(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If


        '2.2.53
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 53)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLPERMISSIONS_BOOLFORCEWATERMARK"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLPERMISSIONS_BOOLFORCEWATERMARK(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLFORCEWATERMARK"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLDATATABLEROWTITLES_BOOLFORCEWATERMARK(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_SIGFIGS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call ModifyRecords_TBLDATATABLEROWTITLES_SIGFIGS(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_CORRECT_CHARCHEMSTRUCTURE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call ModifyRecords_TBLDATATABLEROWTITLES_CORRECT_CHARCHEMSTRUCTURE(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If


        '2.2.56
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 56)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLCONFIGURATION_GOTOWORD"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLCONFIGURATION_GOTOWORD(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If


        '2.2.60
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 60)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLFIELDCODES_LOCKSECTION"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLFIELDCODES_LOCKSECTION(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If


        '2.2.63
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 63)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TableOfContents"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call ModifyRecords_TableOfContents(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If


        '2.2.68
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 68)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "GuWuToStudyDoc_01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call GuWuToStudyDoc_01(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
            'GuWuToStudyDoc_01
        End If


        '2.2.69
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(2, 2, 69)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGREPORTTYPE_OTHER"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLCONFIGREPORTTYPE_OTHER(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call ModifyRecords_TableOfContents_01(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
            'ModifyRecords_TableOfContents_01
        End If


        '3.0.3
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(3, 0, 3)
        If strNV > strNV_DB Then
            'Hmmm. These didn't run correctly. Try again
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TableOfContents"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call ModifyRecords_TableOfContents(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call ModifyRecords_TableOfContents_01(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
        End If


        '3.0.5
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(3, 0, 5)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLREPORTTABLE_BOOLPLACEHOLDER"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLREPORTTABLE_BOOLPLACEHOLDER(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLDATATABLEROWTITLES_BOOLPLACEHOLDER(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLTAB1_PERMISSIONSMANAGER(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLPERMISSIONS_BOOLPERMISSIONS(boolAccess, boolSQLSer, boolOra, con, cmd)
                'Call CreateTable_TBLPERMISSIONSMANAGER(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
                'Call AddRecords_TBLMAXID_TBLPERMISSIONMANAGER(boolAccess, boolSQLSer, boolOra, con, cmd)
                'Call AddRecords_TBLTABLES_TBLPERMISSIONMANAGER(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLUSERACCOUNTS_ID_TBLPERMISSIONS(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call DropFields_TBLPERMISSIONS_ID_USERACCOUNTS(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call DeleteRecords_TBLPERMISSIONS_ID_TBLUSERACCOUNTS(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLPERMISSIONS_CHARPERMISSIONSNAME(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.7
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(3, 0, 7)
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLPERMISSIONS_CHARPERMISSIONSNAME"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call ModifyFields_TBLUSERACCOUNTS_BOOLACTIVE_01(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.8

        '3.0.8
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNV(3, 0, 9)
        tMaxID = 1
        If strNV > strNV_DB Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLWORDDOCS_INTWORDVERSION"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLWORDDOCS_INTWORDVERSION(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call CreateTable_TBLWORDSTATEMENTSVERSIONS(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
                Call AddRecords_TBLTABLES_TBLWORDSTATEMENTSVERSION(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLMAXID_TBLWORDSTATEMENTSVERSION(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLCONFIGURATION_ETMANAGEMENT(boolAccess, boolSQLSer, boolOra, con, cmd)

                Call AddFields_TBLDATA_BOOLPSL(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLDATATABLEROWTITLES_BOOLPSL(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.9.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 9, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLSECTIONTEMPLATES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call CreateTable_TBLSECTIONTEMPLATES(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
                Call AddRecords_TBLMAXID_TBLSECTIONTEMPLATES(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLTABLES_TBLSECTIONTEMPLATES(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.13.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 13, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLSECTIONTEMPLATES"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_tblUserAccounts_dtLogonTime(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.16
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 16, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDROPDOWNBOXCONTENT_01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call ModifyRecords_TBLDROPDOWNBOXCONTENT_01(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLDATA_ROUNDCONV(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If

        '3.0.27
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 27, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATATABLEROWTITLES_01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call ModifyRecords_TBLDATATABLEROWTITLES_01(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call ModifyFields_TBLDATATABLEROWTITLES_TBLCOMPANYDATA_RENUMBER_01(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLDATATABLEROWTITLES_TBLCOMPANYDATA_01(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLDATA_NUMSIGFIGSAREARATIO(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If

        '3.0.27
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 27, 3)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_CHARCAPTIONTRAILER"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLDATATABLEROWTITLES_CHARCAPTIONTRAILER(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLDATA_CHARCAPTIONTRAILER(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLTABLES_TBLLOGIN(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLMAXID_TBLLOGIN(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call CreateTable_TBLLOGIN(boolAccess, boolSQLSer, boolOra, con, cmd, constr)

            Catch ex As Exception

            End Try

        End If

        '3.0.27.4
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 27, 4)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLREPORTTABLEANALYTES_ID_TBLREPORTTABLEANALYTES1"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLREPORTTABLEANALYTES_INTGROUP(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLMAXID_TBLREPORTTABLEANALYTES(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLREPORTTABLEANALYTES_ID_TBLREPORTTABLEANALYTES1(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call ModifyRecords_TBLREPORTTABLEANALYTES_ID_TBLREPORTTABLEANALYTES1(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
                Call AddFields_TBLREPORTTABLEANALYTES_SETPRIMARYKEYS(frmUpdate, boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.27.5
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 27, 5)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLASSIGNEDSAMPLES_INTANALYTEID"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLASSIGNEDSAMPLES_INTANALYTEID(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.27.6
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 27, 6)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_ASSAYDATETIME"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLCONFIGHEADERLOOKUP_ASSAYDATETIME(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.27.7
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 27, 7)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_RUNTYPE"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLCONFIGHEADERLOOKUP_RUNTYPE(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.27.8
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 27, 8)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_AnalRunMatrix"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddRecords_TBLCONFIGHEADERLOOKUP_AnalRunMatrix(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLCONFIGURATION_MDBLOCATION(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.30.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 30, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLWORDSTATEMENTSVERSIONS_IDS"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try
                Call AddFields_TBLWORDSTATEMENTSVERSIONS_IDS(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call CreateTable_TBLFINALREPORT(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
                Call CreateTable_TBLFINALREPORTWORDDOCS(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
                Call AddRecords_TBLTABLES_TBLFINALREPORT(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLMAXID_TBLFINALREPORT(boolAccess, boolSQLSer, boolOra, con, cmd)


                'Call AddRecords_TBLMAXID_TBLFINALREPORT(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.32.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 32, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_AnalRunDate"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try

                Call AddRecords_TBLCONFIGHEADERLOOKUP_AnalRunDate(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.40.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 40, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_DataTableRowTitles_01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try

                Call ModifyRecords_DataTableRowTitles_01(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLDATATABLEROWTITLES_PERMISSIONSREPORTTEMPLATEFINALREPORT(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call ModifyRecords_DataTableRowTitlesSort_01(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.40.3
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 40, 3)
        tMaxID = 1
        If strNV > strNV_DB_New Then
            str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLPERMISSIONS_01"
            frmUpdate.lbl1.Text = str2
            frmUpdate.lbl1.Refresh()
            Try

                Call AddFields_TBLPERMISSIONS_01(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If

        '3.0.40.4
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 40, 4)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLRECSIGFIG"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_BOOLRECSIGFIG(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_BOOLRECSIGFIG"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLDATATABLEROWTITLES_BOOLRECSIGFIG(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_RECOVERYTABLES"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLCONFIGHEADERLOOKUP_RECOVERYTABLES(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLREPORTTABLEHEADERCONFIG_RECOVERYTABLES"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLREPORTTABLEHEADERCONFIG_RECOVERYTABLES(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try
            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLREPORTTABLEHEADERCONFIG_01"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLREPORTTABLEHEADERCONFIG_FIX01(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "ISR_01"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call ISR_01(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "CREATETABLE_TBLAUTOASSIGNSAMPLES"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call CREATETABLE_TBLAUTOASSIGNSAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
            Catch ex As Exception
                var1 = ex.Message
            End Try


            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLMAXID_TBLAUTOASSIGNSAMPLES"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLMAXID_TBLAUTOASSIGNSAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTABLES_TBLAUTOASSIGNSAMPLES"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLTABLES_TBLAUTOASSIGNSAMPLES(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLCONFIGHEADERLOOKUP_Carryover"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLCONFIGHEADERLOOKUP_Carryover(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLAPPFIGWORDDOCS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call CreateTable_TBLAPPFIGWORDDOCS(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLMAXID_TBLAPPFIGWORDDOCS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLMAXID_TBLAPPFIGWORDDOCS(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLTABLES_TBLAPPFIGWORDDOCS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLTABLES_TBLAPPFIGWORDDOCS(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

        End If


        '3.0.42.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 42, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLREPORTS_ANALRUNSUMMARY"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLREPORTS_ANALRUNSUMMARY(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try
           

        End If

        '3.0.48.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 48, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_BOOLINCLMFCOLS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLTABLEPROPERTIES_BOOLINCLMFCOLS(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try


        End If


        '3.0.50.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 51, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            'This needs to be run again because two columns got added to code after v3.0.40
            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_tblUserAccounts_CHARLDAP"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_tblUserAccounts_CHARLDAP(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try


        End If


        '3.0.50.3
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 51, 3)
        tMaxID = 1
        If strNV > strNV_DB_New Then

         Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_CHARLDAP"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLDATATABLEROWTITLES_CHARLDAP(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = ex.Message
            End Try

        End If


        '3.0.53.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 53, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_NUMPRECCRITLOTS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLTABLEPROPERTIES_NUMPRECCRITLOTS(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If

        '3.0.53.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 53, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddRecords_TBLDATATABLEROWTITLES_NUMPRECCRITLOTS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddRecords_TBLDATATABLEROWTITLES_NUMPRECCRITLOTS(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If

        '3.0.55.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 55, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLSD2"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_BOOLSD2(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If


        '3.0.55.3
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 55, 3)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_INTQCLEVELGROUP"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLTABLEPROPERTIES_INTQCLEVELGROUP(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If
        '

        '3.0.58.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 58, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLCUSTOMFIELDCODES_BOOLINCLUDE"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLCUSTOMFIELDCODES_BOOLINCLUDE(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If


        '3.0.61.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 61, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLFINALREPORT_CHARPASSWORD"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLFINALREPORT_CHARPASSWORD(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If


        '3.0.61.3
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 61, 3)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLAPPFIGS_CHARRDB"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLAPPFIGS_CHARRDB(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If


        '3.0.64.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 64, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_CHARSAMPLESAD5"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLTABLEPROPERTIES_CHARSAMPLESAD5(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If


        '3.0.64.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 65, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_CHARLLOQ"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_CHARLLOQ(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLPERMISSIONS_BOOLLOCKFINALREPORT(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If

        '3.0.66.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 66, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_CHARCAPTIONFOLLOW"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_CHARCAPTIONFOLLOW(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If

        '3.0.66.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 66, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "CreateTable_TBLSTUDYDOCANALYTES"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call CreateTable_TBLSTUDYDOCANALYTES(boolAccess, boolSQLSer, boolOra, con, cmd, constr)
                Call AddRecords_TBLMAXID_TBLSTUDYDOCANALYTES(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddRecords_TBLTABLES_TBLSTUDYDOCANALYTES(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If


        '3.0.67.1
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 67, 1)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLUSERSD"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_BOOLUSERSD(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If


        '3.0.68.2
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 68, 2)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLTABLELABELSECTION"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_BOOLTABLELABELSECTION(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLCUSTOMFIELDCODES_INTORDER"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLCUSTOMFIELDCODES_INTORDER(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception

            End Try

        End If


        '3.0.68.3
        'AddFields_TBLTABLEPROPERTIES_BOOLCONCCOMMENTS
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 0, 68, 3)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_BOOLCONCCOMMENTS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLTABLEPROPERTIES_BOOLCONCCOMMENTS(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If


        'AddFields_TBLTABLEPROPERTIES_BOOLADHOCSTABCOMPCOLUMNS
        '3.0.68.3
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 1, 0, 17)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_BOOLADHOCSTABCOMPCOLUMNS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLTABLEPROPERTIES_BOOLADHOCSTABCOMPCOLUMNS(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception

            End Try

        End If


        'AddFields_TBLDATA_NUMTABLEFONTSIZE
        '3.1.0.25
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 1, 0, 25)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_NUMTABLEFONTSIZE"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_NUMTABLEFONTSIZE(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception
                var1 = var1
            End Try

        End If


        'AddFields_TBLSTUDYDOCANALYTES_CHARUSERANALNAME
        '3.1.0.27
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 1, 0, 27)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLSTUDYDOCANALYTES_CHARUSERANALNAME"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLSTUDYDOCANALYTES_CHARUSERANALNAME(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception
                var1 = var1
            End Try

        End If

        'AddFields_TBLAUTOASSIGNSAMPLES_CHARLOTWOIS
        '3.1.0.28
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 1, 0, 28)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLAUTOASSIGNSAMPLES_CHARLOTWOIS"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLAUTOASSIGNSAMPLES_CHARLOTWOIS(boolAccess, boolSQLSer, boolOra, con, cmd)

            Catch ex As Exception
                var1 = var1
            End Try

        End If

        '3.1.0.29
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 1, 0, 29)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLDATA_BOOLCALIBRTABLETITLE"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLDATA_BOOLCALIBRTABLETITLE(boolAccess, boolSQLSer, boolOra, con, cmd)
                Call AddFields_TBLMETHODVALIDATIONDATA_CHARBLOOD(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = var1
            End Try

        End If

        '3.1.0.37
        Call UpdatePG(intPG, intPGMax, frmUpdate, pg)
        strNV = GetNVNew(3, 1, 0, 37)
        tMaxID = 1
        If strNV > strNV_DB_New Then

            Try
                str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLAUTOASSIGNSAMPLES_CHARNEW2"
                frmUpdate.lbl1.Text = str2
                frmUpdate.lbl1.Refresh()
                Call AddFields_TBLAUTOASSIGNSAMPLES_CHARNEW2(boolAccess, boolSQLSer, boolOra, con, cmd)
            Catch ex As Exception
                var1 = var1
            End Try

        End If

        '

        'all versions
        str2 = str1 & ChrW(10) & ChrW(10) & "AddFields_TBLTABLEPROPERTIES_UPDATENULLS"
        frmUpdate.lbl1.Text = str2
        frmUpdate.lbl1.Refresh()
        Call AddFields_TBLTABLEPROPERTIES_UPDATENULLS(boolAccess, boolSQLSer, boolOra, con, cmd)

        str2 = str1 & ChrW(10) & ChrW(10) & "CheckTblReportsNameNull"
        frmUpdate.lbl1.Text = str2
        frmUpdate.lbl1.Refresh()
        Call CheckTblReportsNameNull(boolAccess, boolSQLSer, boolOra, con, cmd)

        str2 = str1 & ChrW(10) & ChrW(10) & "ModifyRecords_TBLDATA_CHARUNITS"
        frmUpdate.lbl1.Text = str2
        frmUpdate.lbl1.Refresh()
        Call ModifyRecords_TBLDATA_CHARUNITS(boolAccess, boolSQLSer, boolOra, con, cmd)

        str2 = str1 & ChrW(10) & ChrW(10) & "ModifyFields_TBLDATATABLEROWTITLES_NUMBERING"
        frmUpdate.lbl1.Text = str2
        frmUpdate.lbl1.Refresh()
        Call ModifyFields_TBLDATATABLEROWTITLES_NUMBERING(boolAccess, boolSQLSer, boolOra, con, cmd)

        str2 = str1 & ChrW(10) & ChrW(10) & "ModifyFields_TBLDATATABLEROWTITLES_NUMBERING_01"
        frmUpdate.lbl1.Text = str2
        frmUpdate.lbl1.Refresh()
        Call ModifyFields_TBLDATATABLEROWTITLES_NUMBERING_01(boolAccess, boolSQLSer, boolOra, con, cmd)
        'ModifyFields_TBLDATATABLEROWTITLES_NUMBERING_01

        frmUpdate.lbl1.Text = str1
        frmUpdate.lbl1.Refresh()

        frmUpdate.pgOverall.Value = frmUpdate.pgOverall.Maximum
        frmUpdate.pgOverall.Refresh()

        var1 = var1

        Cursor.Current = Cursors.Default

    End Sub


    Function GetNV(i1 As Short, i2 As Short, i3 As Short) As String

        GetNV = Format(i1, "00") & Format(i2, "00") & Format(i3, "00")

    End Function

    Function GetNVNew(i1 As Short, i2 As Short, i3 As Short, i4 As Short) As String

        GetNVNew = Format(i1, "00") & Format(i2, "00") & Format(i3, "00") & Format(i4, "00")

    End Function

End Module
