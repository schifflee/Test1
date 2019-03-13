Option Compare Text

Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.IO
Imports System.Text
Imports System.IO.FileSystemInfo

Module modGroups

    'The Group concept creates 'groups' of analytes, matrices, and different calibration sets
    'Think of the set of groups as a database table with a clustered primary key consisting of AnalyteID, IntStd, Matrix, and Calibration Set
    '20160209 LEE: Remove IntStd from consideration at this point
    '   AbbVie has the same IntStd with different names
    '   At some point, give user the choice to differentiate between IntStds
    'The Group concept contains three main tables
    '  - tblAnalyteGroups
    '      a list of each analyte/matrix/calibration set
    '      this table is created in Sub EstablishCalStdGroups
    '          tblAnalyteGroups = dv1a.ToTable("a", True, "ANALYTEDESCRIPTION", "ANALYTEID", "INTSTD", "INTGROUP", "ANALYTEDESCRIPTION_C", "MATRIX", "INTCALSET", "CALIBRSET")

    '  - tblCalStdGroupsAll
    '      examples of each calibration standard calibration set associated with each group for ALL analytical runs
    '      see Sub CreatetblCalStdGroups for a list of columns in this table

    '  - tblCalStdGroupsAcc
    '      examples of each calibration standard calibration set associated with each group for ACCEPTED analytical runs
    '      same columns as tblCalStdGroupsAll

    '  - tblCalStdGroupAssayIDsAll
    '      a list of all the RunIDs (AssayID) associated with each group for ALL analytical runs
    '      see Sub CreatetblCalStdGroupAssayIDs for a list of columns in this table

    '  - tblCalStdGroupAssayIDsAcc
    '      a list of all the RunIDs (AssayID) associated with each group for ALL analytical runs
    '      same columns as tblCalStdGroupAssayIDsAll

    '20150912 Larry: New Group paradigm
    Public boolUseGroups As Boolean = False
    Public tblCalStdGroupsAll As New System.Data.DataTable
    Public tblCalStdGroupsAcc As New System.Data.DataTable
    Public tblCalStdGroupAssayIDsAll As New System.Data.DataTable
    Public tblCalStdGroupAssayIDsAcc As New System.Data.DataTable
    Public tblAnalyteGroupsAcc As New System.Data.DataTable
    Public tblAnalyteGroupsTemp As New System.Data.DataTable
    Public tblAnalyteGroups As New System.Data.DataTable


    Sub CreatetblCalStdGroups()


        'this table lists the calibration levels of each Group

        If tblCalStdGroupsAll.Columns.Contains("ANALYTEDESCRIPTION") Then
            Exit Sub
        End If

        Dim str1 As String

        Dim col4 As New DataColumn
        str1 = "ANALYTEDESCRIPTION"
        col4.DataType = System.Type.GetType("System.String")
        col4.ColumnName = str1
        col4.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col4)

        Dim col4a As New DataColumn
        str1 = "ANALYTEDESCRIPTION_C"
        col4a.DataType = System.Type.GetType("System.String")
        col4a.ColumnName = str1
        col4a.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col4a)

        Dim col5 As New DataColumn
        str1 = "INTSTD"
        col5.DataType = System.Type.GetType("System.String")
        col5.ColumnName = str1
        col5.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col5)

        Dim col6a As New DataColumn
        str1 = "MATRIX"
        col6a.DataType = System.Type.GetType("System.String")
        col6a.ColumnName = str1
        col6a.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col6a)

        Dim col7 As New DataColumn
        str1 = "ANALYTEID"
        col7.DataType = System.Type.GetType("System.Int64")
        col7.ColumnName = str1
        col7.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col7)

        Dim col8 As New DataColumn
        str1 = "ANALYTEINDEX"
        col8.DataType = System.Type.GetType("System.Int64")
        col8.ColumnName = str1
        col8.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col8)

        Dim col2 As New DataColumn
        str1 = "LEVELNUMBER"
        col2.DataType = System.Type.GetType("System.Int16")
        col2.ColumnName = str1
        col2.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col2)

        Dim col3 As New DataColumn
        str1 = "CONCENTRATION"
        col3.DataType = System.Type.GetType("System.Single")
        col3.ColumnName = str1
        col3.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col3)

        Dim col10 As New DataColumn
        str1 = "RUNID"
        col10.DataType = System.Type.GetType("System.Int16")
        col10.ColumnName = str1
        col10.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col10)

        Dim col10A As New DataColumn
        str1 = "ASSAYID"
        col10A.DataType = System.Type.GetType("System.Int64")
        col10A.ColumnName = str1
        col10A.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col10A)

        Dim col10b As New DataColumn
        str1 = "MASTERASSAYID"
        col10b.DataType = System.Type.GetType("System.Int64")
        col10b.ColumnName = str1
        col10b.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col10b)

        Dim col6c As New DataColumn
        str1 = "ANALYTEFLAGPERCENT"
        col6c.DataType = System.Type.GetType("System.Single")
        col6c.ColumnName = str1
        col6c.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col6c)

        Dim col11 As New DataColumn
        str1 = "RUNDATE"
        col11.DataType = System.Type.GetType("System.DateTime")
        col11.ColumnName = str1
        col11.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col11)

        Dim col6b As New DataColumn
        str1 = "RUNTYPEID"
        col6b.DataType = System.Type.GetType("System.Int16")
        col6b.ColumnName = str1
        tblCalStdGroupsAll.Columns.Add(col6b)

        Dim col6ba As New DataColumn
        str1 = "RUNTYPE"
        col6ba.DataType = System.Type.GetType("System.String")
        col6ba.ColumnName = str1
        tblCalStdGroupsAll.Columns.Add(col6ba)


        Dim col6bb As New DataColumn
        str1 = "RUNANALYTEREGRESSIONSTATUS"
        col6bb.DataType = System.Type.GetType("System.Int16")
        col6bb.ColumnName = str1
        tblCalStdGroupsAll.Columns.Add(col6bb)


        Dim col6bc As New DataColumn
        str1 = "LLOQ"
        col6bc.DataType = System.Type.GetType("System.Single")
        col6bc.ColumnName = str1
        tblCalStdGroupsAll.Columns.Add(col6bc)

        Dim col6bd As New DataColumn
        str1 = "ULOQ"
        col6bd.DataType = System.Type.GetType("System.Single")
        col6bd.ColumnName = str1
        tblCalStdGroupsAll.Columns.Add(col6bd)

        Dim col6be As New DataColumn
        str1 = "CONCENTRATIONUNITS"
        col6be.DataType = System.Type.GetType("System.String")
        col6be.ColumnName = str1
        col6be.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col6be)

        Dim col6d As New DataColumn
        str1 = "INTCALSET"
        col6d.DataType = System.Type.GetType("System.Int16")
        col6d.ColumnName = str1
        col6d.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col6d)

        Dim col6 As New DataColumn
        str1 = "INTGROUP"
        col6.DataType = System.Type.GetType("System.Int16")
        col6.ColumnName = str1
        col6.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col6)

        Dim col6e As New DataColumn
        str1 = "CALIBRSET"
        col6e.DataType = System.Type.GetType("System.String")
        col6e.ColumnName = str1
        col6e.Caption = str1 'strAnal
        tblCalStdGroupsAll.Columns.Add(col6e)



    End Sub

    Sub CreatetblCalStdGroupAssayIDs()

        If tblCalStdGroupAssayIDsAll.Columns.Contains("ANALYTEDESCRIPTION") Then
            Exit Sub
        End If

        Dim str1 As String

        Dim col4a As New DataColumn
        str1 = "ANALYTEDESCRIPTION"
        col4a.DataType = System.Type.GetType("System.String")
        col4a.ColumnName = str1
        col4a.Caption = str1 'strAnal

        Dim col4b As New DataColumn
        str1 = "ANALYTEDESCRIPTION_C"
        col4b.DataType = System.Type.GetType("System.String")
        col4b.ColumnName = str1
        col4b.Caption = str1 'strAnal

        Dim col5a As New DataColumn
        str1 = "INTSTD"
        col5a.DataType = System.Type.GetType("System.String")
        col5a.ColumnName = str1
        col5a.Caption = str1 'strAnal

        Dim col6a As New DataColumn
        str1 = "MATRIX"
        col6a.DataType = System.Type.GetType("System.String")
        col6a.ColumnName = str1
        col6a.Caption = str1 'strAnal

        Dim col7a As New DataColumn
        str1 = "ANALYTEID"
        col7a.DataType = System.Type.GetType("System.Int64")
        col7a.ColumnName = str1
        col7a.Caption = str1 'strAnal

        Dim col8a As New DataColumn
        str1 = "ANALYTEINDEX"
        col8a.DataType = System.Type.GetType("System.Int64")
        col8a.ColumnName = str1
        col8a.Caption = str1 'strAnal

        Dim col1a As New DataColumn
        str1 = "ASSAYID"
        col1a.DataType = System.Type.GetType("System.Int64")
        col1a.ColumnName = str1
        col1a.Caption = str1 'strAnal

        Dim col9a As New DataColumn 'this parameter probably not needed, but record it anyway just in case
        str1 = "MASTERASSAYID"
        col9a.DataType = System.Type.GetType("System.Int64")
        col9a.ColumnName = str1
        col9a.Caption = str1 'strAnal

        Dim col10a As New DataColumn
        str1 = "RUNID"
        col10a.DataType = System.Type.GetType("System.Int16")
        col10a.ColumnName = str1
        col10a.Caption = str1 'strAnal

        Dim col11a As New DataColumn
        str1 = "RUNDATE"
        col11a.DataType = System.Type.GetType("System.DateTime")
        col11a.ColumnName = str1
        col11a.Caption = str1 'strAnal

        Dim col6b As New DataColumn
        str1 = "RUNTYPEID"
        col6b.DataType = System.Type.GetType("System.Int16")
        col6b.ColumnName = str1
        col6b.Caption = str1 'strAnal

        Dim col6ba As New DataColumn
        str1 = "RUNTYPE"
        col6ba.DataType = System.Type.GetType("System.String")
        col6ba.ColumnName = str1
        col6ba.Caption = str1 'strAnal

        Dim col6bb As New DataColumn
        str1 = "RUNANALYTEREGRESSIONSTATUS"
        col6bb.DataType = System.Type.GetType("System.Int16")
        col6bb.ColumnName = str1

        Dim col6bc As New DataColumn
        str1 = "LLOQ"
        col6bc.DataType = System.Type.GetType("System.Single")
        col6bc.ColumnName = str1

        Dim col6bd As New DataColumn
        str1 = "ULOQ"
        col6bd.DataType = System.Type.GetType("System.Single")
        col6bd.ColumnName = str1

        Dim col6be As New DataColumn
        str1 = "CONCENTRATIONUNITS"
        col6be.DataType = System.Type.GetType("System.String")
        col6be.ColumnName = str1
        col6be.Caption = str1 'strAnal

        Dim col6d As New DataColumn
        str1 = "INTCALSET"
        col6d.DataType = System.Type.GetType("System.Int16")
        col6d.ColumnName = str1
        col6d.Caption = str1 'strAnal

        Dim col6c As New DataColumn
        str1 = "INTGROUP"
        col6c.DataType = System.Type.GetType("System.Int16")
        col6c.ColumnName = str1
        col6c.Caption = str1 'strAnal

        tblCalStdGroupAssayIDsAll.Columns.Add(col4a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col4b)
        tblCalStdGroupAssayIDsAll.Columns.Add(col5a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col6a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col7a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col8a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col1a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col9a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col10a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col11a)
        tblCalStdGroupAssayIDsAll.Columns.Add(col6b)
        tblCalStdGroupAssayIDsAll.Columns.Add(col6ba)
        tblCalStdGroupAssayIDsAll.Columns.Add(col6bb)

        tblCalStdGroupAssayIDsAll.Columns.Add(col6bc)
        tblCalStdGroupAssayIDsAll.Columns.Add(col6bd)
        tblCalStdGroupAssayIDsAll.Columns.Add(col6be)

        tblCalStdGroupAssayIDsAll.Columns.Add(col6d)
        tblCalStdGroupAssayIDsAll.Columns.Add(col6c)

    End Sub

    Sub EstablishCalStdGroups(boolStd As Boolean, con As ADODB.Connection)

        ''debug
        ''tblAnalytesHome
        'Dim CC1 As Integer
        'Dim CC2 As Integer
        'Dim v1, v2
        'Dim tblA As DataTable = tblAnalytesHome

        'v1 = ""
        'For CC1 = 0 To tblA.Columns.Count - 1
        '    v2 = tblA.Columns(CC1).ColumnName
        '    v1 = v1 & ";" & v2
        'Next CC1
        'Console.WriteLine(v1)

        'For CC2 = 0 To tblA.Rows.Count - 1
        '    v1 = ""
        '    For CC1 = 0 To tblA.Columns.Count - 1
        '        v2 = tblA.Rows(CC2).Item(CC1)
        '        v1 = v1 & ";" & v2
        '    Next CC1
        '    Console.WriteLine(v1)
        'Next


        'this routine requires the following tables to be populated beforehand

        'tblBCStdsAssayID
        'tblBCQCStdsAssayID

        'first clear existing tables
        'tblCalStdGroupsAll
        'tblCalStdGroupAssayIDs

        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9, var10, var11, var12
        Dim Count1 As Int32
        Dim int1 As Int32

        'first clear the contents of existing datatables
        'Note: cannot use .clear
        '  .clear only removes rows, it doesn't delete them
        'must use a loop

        int1 = tblCalStdGroupsAll.Rows.Count
        For Count1 = int1 To 1 Step -1
            tblCalStdGroupsAll.Rows(Count1 - 1).Delete()
        Next

        int1 = tblCalStdGroupsAcc.Rows.Count
        For Count1 = int1 To 1 Step -1
            tblCalStdGroupsAcc.Rows(Count1 - 1).Delete()
        Next

        int1 = tblCalStdGroupAssayIDsAll.Rows.Count
        For Count1 = int1 To 1 Step -1
            tblCalStdGroupAssayIDsAll.Rows(Count1 - 1).Delete()
        Next

        int1 = tblCalStdGroupAssayIDsAcc.Rows.Count
        For Count1 = int1 To 1 Step -1
            tblCalStdGroupAssayIDsAcc.Rows(Count1 - 1).Delete()
        Next

        Try
            int1 = tblAnalyteGroupsTemp.Rows.Count
            For Count1 = int1 To 1 Step -1
                tblAnalyteGroupsTemp.Rows(Count1 - 1).Delete()
            Next
        Catch ex As Exception

        End Try

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String
        Dim strSQL As String

        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim Count4 As Int32
        Dim Count5 As Int32
        Dim Count6 As Int32
        Dim Count7 As Int32
        Dim Count8 As Int32

        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim strF4 As String
        Dim strF5 As String
        Dim strF6 As String
        Dim strF7 As String
        Dim strF8 As String
        Dim strF9 As String
        Dim strF10 As String
        Dim strS As String

        Dim vRunID, vDate


        Dim int2 As Int32
        Dim int3 As Int32
        Dim int4 As Int32
        Dim strAnalDescr As String
        Dim strIntStd As String
        Dim intGroup As Int32
        Dim boolHit As Boolean
        Dim boolHit1 As Boolean
        Dim boolHit2 As Boolean

        Dim intL1 As Short
        Dim intL2 As Short

        Dim numConc1 As Single
        Dim numConc2 As Single

        Dim c1 As Single
        Dim c2 As Single
        Dim l1 As Short
        Dim l2 As Short

        Dim dtNow As Date = Now

        Dim strSampleType As String
        Dim intAssayID As Int64
        Dim intAnalyteIndex As Int64
        Dim intAnalyteID As Int64
        Dim intMasterAssayID As Int64
        Dim strAnalyteName As String
        Dim strIntStdName As String
        Dim intLevelNumber As Int16
        Dim numNomConc As Single
        Dim intRunID As Int16
        Dim dtRunDate As Date
        Dim intRunTypeID As Short

        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim rows3() As DataRow
        Dim rows4() As DataRow
        Dim rows5() As DataRow
        Dim rows6() As DataRow

        Dim boolDiffIS As Boolean = False '

        Dim intGroupA As Int32

        Dim boolCoCmpd As Boolean
        Dim strCoCmpd As String

        Dim intR As Integer
        Dim strM As String

        Dim tblStdGroups As System.Data.DataTable
        Dim tblStdGroupAssayIDs As System.Data.DataTable
        Dim tblStdAssayID As System.Data.DataTable

        Dim tblAssayIntStd As New System.Data.DataTable
        Dim rowIntStd() As DataRow

        If boolAccess Then
            str1 = "SELECT ASSAYANALYTES.* FROM ASSAYANALYTES "
            str2 = "WHERE (((ASSAYANALYTES.STUDYID)=" & wStudyID & "));"

            str1 = "SELECT CONFIGSAMPLETYPES.SAMPLETYPEID, ASSAYANALYTES.* "
            str2 = "FROM ASSAYANALYTES INNER JOIN (ASSAY INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (ASSAYANALYTES.ASSAYID = ASSAY.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAY.STUDYID) "
            str3 = "WHERE (((ASSAYANALYTES.STUDYID)=" & wStudyID & "));"

        Else
            str1 = "SELECT " & strSchema & ".ASSAYANALYTES.* FROM " & strSchema & ".ASSAYANALYTES "
            str2 = "WHERE (((" & strSchema & ".ASSAYANALYTES.STUDYID)=" & wStudyID & "));"

            str1 = "SELECT " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".ASSAYANALYTES.* "
            str2 = "FROM " & strSchema & ".ASSAYANALYTES INNER JOIN (" & strSchema & ".ASSAY INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAY.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAY.STUDYID) "
            str3 = "WHERE (((" & strSchema & ".ASSAYANALYTES.STUDYID)=" & wStudyID & "));"

        End If

        strSQL = str1 & str2 & str3

        '''Console.WriteLine("tblAssayIntStd: " & strSQL)

        Dim rs As New ADODB.Recordset
        rs.CursorLocation = CursorLocationEnum.adUseClient
        rs.Open(strSQL, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
        rs.ActiveConnection = Nothing

        Dim da As OleDbDataAdapter = New OleDbDataAdapter()

        tblAssayIntStd.Clear()
        tblAssayIntStd.AcceptChanges()
        tblAssayIntStd.BeginLoadData()
        da.Fill(tblAssayIntStd, rs)
        tblAssayIntStd.EndLoadData()
        rs.Close()
        rs = Nothing


        Dim numAFP As Single 'analyteflagpercent

        'Group priority:
        '1. Matrix
        '2. Calibr Levels
        '3. Analyte

        boolStd = True
        strSampleType = "Standard"
        'tblStdGroups = tblCalStdGroupsAll
        'tblStdGroupAssayIDs = tblCalStdGroupAssayIDsAll
        '20160204 LEE: tblAllStdsAssay has all analytical runs
        strF = "KNOWNTYPE='STANDARD'"
        Dim dvASA As DataView = New DataView(tblAllStdsAssay, strF, "", DataViewRowState.CurrentRows)
        tblStdAssayID = dvASA.ToTable ' tblAllStdsAssay 'tblBCStdsAssayID
        'tblStdAssayID = tblBCStdsAssayID

        'start with accepted analytical runs
        'no, do ALL anal runs
        Dim tblAAR As System.Data.DataTable = tblAllAnalRuns ' tblAccAnalRuns
        Dim intNumAAR As Int16 = tblAAR.Rows.Count

        intGroup = 0

        Dim CountM As Short
        Dim strMatrix As String
        Dim intCalSet As Short
        Dim intCalSetR As Short
        Dim gCalSet As Short

        Dim strRunType As String

        Dim CountID As Short
        strF = "RUNID > 0"
        strS = "ANALYTEID ASC"
        Dim dvID As DataView = New DataView(tblAllAnalRuns, strF, strS, DataViewRowState.CurrentRows)
        Dim tblAnalIDs As DataTable
        Try
            tblAnalIDs = dvID.ToTable("b", True, "ANALYTEID", "ANALYTEDESCRIPTION")
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Dim numAnalIDs As Short = tblAnalIDs.Rows.Count
        Dim intRARS As Int64

        Dim rowAQL() As DataRow
        Dim numLLOQ As Single
        Dim numULOQ As Single
        Dim strConcUnits As String

        Dim intGroupD As Int32

        '20171111 LEE:
        'legend from doPrepare
        'Dim dvSP As System.Data.DataView = New DataView(tblSpeciesMatrix)
        'Dim tblSP As System.Data.DataTable = dvSP.ToTable("aSP", True, "SPECIES")
        'intNumSpecies = tblSP.Rows.Count
        ''find number of matrixes
        'tblSP = dvSP.ToTable("aSM", True, "SAMPLETYPEID")
        'intNumMatrix = tblSP.Rows.Count
        'gNumMatrix = intNumMatrix

        'tblSpeciesMatrix can have replicate matrices because SampleVolume may have different entries
        'so can't use tblspeciesmatrix
        '
        Dim dvSP As System.Data.DataView = New DataView(tblSpeciesMatrix)
        'must use a unique table
        Dim tblSP As System.Data.DataTable = dvSP.ToTable("aSM", True, "SAMPLETYPEID")
        Dim intNumMatrixHere As Short = tblSP.Rows.Count

        Dim rows2a() As DataRow
        Dim rows2b() As DataRow
        Dim boolSkip1 As Boolean

        ''20171220 LEE:
        ''Alturas SGN19A-001
        ''Assay table has only Plasma
        ''DesignSample table has Plasma and CFS
        ''try getting sampletypeid from tbldesignsample
        'Dim dvSP As System.Data.DataView = New DataView(tblSampleDesign)
        ''must use a unique table
        'Dim tblSP As System.Data.DataTable = dvSP.ToTable("aSM", True, "SAMPLETYPEID")
        'Dim intNumMatrixHere As Short = tblSP.Rows.Count


        'loop through each matrix
        gCalSet = 1
        For CountM = 1 To intNumMatrixHere

            strMatrix = tblSP.Rows(CountM - 1).Item("SAMPLETYPEID")
            var1 = var1 'debug

            'then go through analyteid
            For CountID = 1 To numAnalIDs

                intCalSet = 1

                intAnalyteID = tblAnalIDs.Rows(CountID - 1).Item("ANALYTEID")
                strAnalyteName = tblAnalIDs.Rows(CountID - 1).Item("ANALYTEDESCRIPTION")

                'get unique assayid's from this table
                'set to dataview

                strF1 = "ANALYTEID = " & intAnalyteID
                '20180219 LEE:
                strF1 = "ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                strS = "RUNID ASC"
                Dim dvUAssID As DataView = New DataView(tblAAR, strF1, strS, DataViewRowState.CurrentRows)
                'Dim tblUAID As System.Data.DataTable = dv1.ToTable("a", True, "ASSAYID")
                Dim intNumUAID As Int16 = dvUAssID.Count ' tblUAID.Rows.Count

                Try
                    For Count2 = 0 To intNumUAID - 1 'loop through AssayIDs/RunIDs

                        If Count2 = 6 Then
                            var1 = var1
                        End If

                        intCalSetR = 0

                        intAnalyteIndex = dvUAssID(Count2).Item("ANALYTEINDEX")
                        intMasterAssayID = dvUAssID(Count2).Item("MASTERASSAYID") 'probably don't need this, but record anyway
                        intRunID = dvUAssID(Count2).Item("RUNID")
                        intRunTypeID = dvUAssID(Count2).Item("RUNTYPEID")
                        intAssayID = dvUAssID(Count2).Item("ASSAYID")
                        intRARS = dvUAssID(Count2).Item("RUNANALYTEREGRESSIONSTATUS")
                        strRunType = dvUAssID(Count2).Item("RUNTYPEDESCRIPTION")
                        var1 = dvUAssID(Count2).Item("RUNSTARTDATE")
                        If IsDate(var1) Then
                            dtRunDate = dvUAssID(Count2).Item("RUNSTARTDATE")
                        Else
                            dtRunDate = CDate("01-01-1900")
                        End If


                        'get LLOQ, ULOQ, Conc Units
                        'strF = "RUNID = " & intRunID & " AND ANALYTEID = " & intAnalyteID
                        'rowAQL = tblAllAnalRuns.Select(strF)
                        'numLLOQ = rowAQL(0).Item("NM")
                        'numULOQ = rowAQL(0).Item("VEC")
                        'strConcUnits = rowAQL(0).Item("CONCENTRATIONUNITS")

                        numLLOQ = dvUAssID(Count2).Item("NM")
                        numULOQ = dvUAssID(Count2).Item("VEC")
                        strConcUnits = dvUAssID(Count2).Item("CONCENTRATIONUNITS")

                        'get strIntStdName
                        strF = "ASSAYID = " & intAssayID & " AND ANALYTEID = " & intAnalyteID
                        '20180219 LEE:
                        strF = "ASSAYID = " & intAssayID & " AND ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                        rowIntStd = tblAssayIntStd.Select(strF)
                        If rowIntStd.Length = 0 Then
                            strIntStdName = "IS"
                        Else
                            strIntStdName = NZ(rowIntStd(0).Item("INTERNALSTANDARD"), "NA")
                        End If


                        ''get run date
                        'strF4 = "ASSAYID = " & intAssayID
                        'rows5 = tblAAUnkRunID.Select(strF4)
                        'Dim intDtRows As Short
                        'intDtRows = rows5.Length
                        'dtRunDate = rows5(0).Item("RUNSTARTDATE")

                        'now get the analyteid for this index and assayid
                        'strF2 = "ANALYTEINDEX = " & intAnalyteIndex & " AND ASSAYID = " & intAssayID & " AND SAMPLETYPEID = '" & strMatrix & "'"

                        '20180316 LEE: PROBLEM!
                        'The previous logic works if analytes are represented in runs with OR without calibr standards, but not with AND without calibr stds
                        'Examples to test: 
                        'Alturas - POP01: with OR without calibr standards
                        'Intervet NJ - S1632000: with AND without calibr stds
                        'must look for both assayid (rows2a) and not assayid (rows2b)
                        'if both .length = 0 then continue
                        'elseif 2a.length = 0 and 2b.length > 0 then skip

                        strF2 = "ANALYTEID = " & intAnalyteID & " AND ASSAYID = " & intAssayID & " AND SAMPLETYPEID = '" & strMatrix & "'"
                        strF3 = "ANALYTEID = " & intAnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "'"

                        'this array will return a set of Standard Levels
                        'rows2 = tblStdAssayID.Select(strF2, "ANALYTEID ASC, ANALYTEINDEX ASC, LEVELNUMBER ASC")

                        rows2 = tblStdAssayID.Select(strF2, "ANALYTEID ASC, ANALYTEINDEX ASC, CONCENTRATION ASC")
                        rows2a = tblStdAssayID.Select(strF3, "ANALYTEID ASC, ANALYTEINDEX ASC, CONCENTRATION ASC")

                        boolSkip1 = False
                        If rows2.Length = 0 Then
                            If rows2a.Length = 0 Then
                            Else
                                boolSkip1 = True
                            End If
                        End If

                        '20171219 LEE: PROBLEM! This filter excludes analytical runs that do not have calibration standards! BAD!
                        'do not goto NextCount2 because every analytical run must be accounted for


                        If rows2.Length = 0 Then
                            var1 = var1

                            '20171219 LEE: tblCalStdGroupAssayIDsAll used to exclude runs with no calibr curve
                            'modGroups was modified to ensure tblCalStdGroupAssayIDsAll has ALL analytical runs
                            'do not goto anymore, see above note for explanation


                            '20180316 LEE: PROBLEM!
                            'The previous logic works if analytes are represented in runs with OR without calibr standards, but not with AND without calibr stds
                            'Examples to test: 
                            'Alturas - POP01: with OR without calibr standards
                            'Intervet NJ - S1632000: with AND without calibr stds
                            '20181316 LEE:
                            'evaluate bookSkip1
                            If boolSkip1 Then
                                GoTo NextCount2
                            End If
                            '
                        End If

                        'get analyte name and Int Std from tblAnalyteHome
                        strF3 = "ANALYTEID = " & intAnalyteID
                        '20180219 LEE:
                        strF3 = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                        rows3 = tblAnalytesHome.Select(strF3)
                        If rows3.Length = 0 Then
                            strCoCmpd = "No"
                        Else
                            strCoCmpd = NZ(rows3(0).Item("IsCoadminCmpd"), "No")
                            If StrComp(strCoCmpd, "Yes", CompareMethod.Text) = 0 Then
                                GoTo nextcount2
                            End If
                        End If


                        'record this entry in tblCalStdGroupsAll

                        If intGroup = 0 Then

                            intGroup = intGroup + 1
                            intGroupA = intGroup

                            If rows2.Length = 0 Then

                                '20171219 LEE

                                intLevelNumber = -1
                                numNomConc = -1
                                numAFP = 15

                                Dim nr As DataRow = tblCalStdGroupsAll.NewRow
                                nr.BeginEdit()
                                nr.Item("ANALYTEDESCRIPTION") = strAnalyteName
                                nr.Item("INTSTD") = strIntStdName
                                nr.Item("ANALYTEID") = intAnalyteID
                                nr.Item("ANALYTEINDEX") = intAnalyteIndex
                                nr.Item("ASSAYID") = intAssayID
                                nr.Item("MASTERASSAYID") = intMasterAssayID
                                nr.Item("LEVELNUMBER") = intLevelNumber
                                nr.Item("CONCENTRATION") = numNomConc
                                nr.Item("ANALYTEFLAGPERCENT") = numAFP
                                nr.Item("INTGROUP") = intGroup
                                nr.Item("RUNTYPEID") = intRunTypeID
                                nr.Item("RUNANALYTEREGRESSIONSTATUS") = intRARS

                                nr.Item("RUNTYPE") = strRunType

                                nr.Item("LLOQ") = numLLOQ
                                nr.Item("ULOQ") = numULOQ
                                nr.Item("CONCENTRATIONUNITS") = strConcUnits


                                nr.Item("RUNID") = intRunID
                                nr.Item("RUNDATE") = dtRunDate
                                nr.Item("MATRIX") = strMatrix
                                nr.Item("INTCALSET") = intCalSet

                                nr.EndEdit()

                                Try
                                    tblCalStdGroupsAll.Rows.Add(nr)
                                Catch ex As Exception
                                    var1 = ex.Message
                                End Try
                            Else
                                For Count3 = 0 To rows2.Length - 1

                                    intLevelNumber = NZ(rows2(Count3).Item("LEVELNUMBER"), -1)
                                    numNomConc = NZ(rows2(Count3).Item("CONCENTRATION"), -1)
                                    numAFP = NZ(rows2(Count3).Item("ANALYTEFLAGPERCENT"), NZ(rows2(Count3).Item("FLAGPERCENT"), 15))

                                    Dim nr As DataRow = tblCalStdGroupsAll.NewRow
                                    nr.BeginEdit()
                                    nr.Item("ANALYTEDESCRIPTION") = strAnalyteName
                                    nr.Item("INTSTD") = strIntStdName
                                    nr.Item("ANALYTEID") = intAnalyteID
                                    nr.Item("ANALYTEINDEX") = intAnalyteIndex
                                    nr.Item("ASSAYID") = intAssayID
                                    nr.Item("MASTERASSAYID") = intMasterAssayID
                                    nr.Item("LEVELNUMBER") = intLevelNumber
                                    nr.Item("CONCENTRATION") = numNomConc
                                    nr.Item("ANALYTEFLAGPERCENT") = numAFP
                                    nr.Item("INTGROUP") = intGroup
                                    nr.Item("RUNTYPEID") = intRunTypeID
                                    nr.Item("RUNANALYTEREGRESSIONSTATUS") = intRARS

                                    nr.Item("RUNTYPE") = strRunType

                                    nr.Item("LLOQ") = numLLOQ
                                    nr.Item("ULOQ") = numULOQ
                                    nr.Item("CONCENTRATIONUNITS") = strConcUnits


                                    nr.Item("RUNID") = intRunID
                                    nr.Item("RUNDATE") = dtRunDate
                                    nr.Item("MATRIX") = strMatrix
                                    nr.Item("INTCALSET") = intCalSet

                                    nr.EndEdit()

                                    Try
                                        tblCalStdGroupsAll.Rows.Add(nr)
                                    Catch ex As Exception
                                        var1 = ex.Message
                                    End Try

                                Next Count3

                            End If

                            intCalSetR = intCalSet

                        Else

                            'see if nomconc set already exists in tblCalStdGroupsAll

                            'evaluate for different group
                            'check to see if nom concs are different
                            Dim intGN As Integer

                            'first determine if any group sets have the same number of levels as dv3 AND has the same analyteid

                            Dim arrGG()
                            'ReDim arrGG(intGroup) 
                            ReDim arrGG(intGroup * intCalSet) '20190207 LEE: need to account for intCalSet. Later errors out Frontage 10182592 Sample Analysis
                            Dim intGG As Short = 0
                            For Count5 = 1 To intGroup

                                'strF5 = "INTGROUP = " & Count5
                                'strF5 = "INTGROUP = " & Count5 & " AND ANALYTEID = " & intAnalyteID
                                'strF5 = "INTGROUP = " & Count5 & " AND ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                                'strF5 = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                                strF5 = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "' AND INTCALSET = " & intCalSet
                                ' ''20181122 LEE:
                                ' ''intCalSet should be Count5
                                '20181219 LEE:
                                'No it shouldn't. Count5 is Group, not CalSet
                                'strF5 = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "' AND INTCALSET = " & Count5
                                '20181220 LEE:
                                'Frontage study BTM-2421: This logic is assigning analytes to incorrect run
                                'need to include loop for intCalSet
                              
                                For Count6 = 1 To intCalSet

                                    strF5 = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "' AND INTCALSET = " & Count6

                                    rows4 = tblCalStdGroupsAll.Select(strF5)

                                    If rows4.Length = 0 Then

                                        'make new group
                                        intGroup = intGroup + 1
                                        intGroupA = intGroup

                                        gNumCalSets = gNumCalSets + 1

                                        If rows2.Length = 0 Then

                                            '20171219 LEE

                                            intLevelNumber = -1
                                            numNomConc = -1
                                            numAFP = 15

                                            Dim nr As DataRow = tblCalStdGroupsAll.NewRow
                                            nr.BeginEdit()
                                            nr.Item("ANALYTEDESCRIPTION") = strAnalyteName
                                            nr.Item("INTSTD") = strIntStdName
                                            nr.Item("ANALYTEID") = intAnalyteID
                                            nr.Item("ANALYTEINDEX") = intAnalyteIndex
                                            nr.Item("ASSAYID") = intAssayID
                                            nr.Item("MASTERASSAYID") = intMasterAssayID
                                            nr.Item("LEVELNUMBER") = intLevelNumber
                                            nr.Item("CONCENTRATION") = numNomConc

                                            nr.Item("ANALYTEFLAGPERCENT") = numAFP

                                            nr.Item("INTGROUP") = intGroup
                                            nr.Item("RUNTYPEID") = intRunTypeID
                                            nr.Item("RUNANALYTEREGRESSIONSTATUS") = intRARS

                                            nr.Item("RUNTYPE") = strRunType

                                            nr.Item("LLOQ") = numLLOQ
                                            nr.Item("ULOQ") = numULOQ
                                            nr.Item("CONCENTRATIONUNITS") = strConcUnits

                                            nr.Item("RUNID") = intRunID
                                            nr.Item("RUNDATE") = dtRunDate
                                            nr.Item("MATRIX") = strMatrix
                                            nr.Item("INTCALSET") = Count6 ' intCalSet

                                            nr.EndEdit()

                                            Try
                                                tblCalStdGroupsAll.Rows.Add(nr)
                                            Catch ex As Exception
                                                var1 = ex.Message
                                            End Try

                                        Else

                                            For Count3 = 0 To rows2.Length - 1

                                                intLevelNumber = NZ(rows2(Count3).Item("LEVELNUMBER"), -1)
                                                numNomConc = NZ(rows2(Count3).Item("CONCENTRATION"), -1)
                                                numAFP = NZ(rows2(Count3).Item("ANALYTEFLAGPERCENT"), NZ(rows2(Count3).Item("FLAGPERCENT"), 15))

                                                Dim nr As DataRow = tblCalStdGroupsAll.NewRow
                                                nr.BeginEdit()
                                                nr.Item("ANALYTEDESCRIPTION") = strAnalyteName
                                                nr.Item("INTSTD") = strIntStdName
                                                nr.Item("ANALYTEID") = intAnalyteID
                                                nr.Item("ANALYTEINDEX") = intAnalyteIndex
                                                nr.Item("ASSAYID") = intAssayID
                                                nr.Item("MASTERASSAYID") = intMasterAssayID
                                                nr.Item("LEVELNUMBER") = intLevelNumber
                                                nr.Item("CONCENTRATION") = numNomConc

                                                nr.Item("ANALYTEFLAGPERCENT") = numAFP

                                                nr.Item("INTGROUP") = intGroup
                                                nr.Item("RUNTYPEID") = intRunTypeID
                                                nr.Item("RUNANALYTEREGRESSIONSTATUS") = intRARS

                                                nr.Item("RUNTYPE") = strRunType

                                                nr.Item("LLOQ") = numLLOQ
                                                nr.Item("ULOQ") = numULOQ
                                                nr.Item("CONCENTRATIONUNITS") = strConcUnits

                                                nr.Item("RUNID") = intRunID
                                                nr.Item("RUNDATE") = dtRunDate
                                                nr.Item("MATRIX") = strMatrix
                                                nr.Item("INTCALSET") = Count6 'intCalSet

                                                nr.EndEdit()

                                                Try
                                                    tblCalStdGroupsAll.Rows.Add(nr)
                                                Catch ex As Exception
                                                    var1 = ex.Message
                                                End Try

                                            Next Count3

                                        End If


                                        intCalSetR = intCalSet

                                    ElseIf rows4.Length = rows2.Length Then
                                        intGG = intGG + 1
                                        If intGG > UBound(arrGG) Then '20190207 LEE: added as check
                                            ReDim Preserve arrGG(intGG)
                                        End If
                                        arrGG(intGG) = Count6
                                    Else
                                        ''intGG = intGG + 1
                                        ''arrGG(intGG) = Count5
                                        If intGroup = 1 Then
                                            intGG = intGG + 1
                                            If intGG > UBound(arrGG) Then '20190207 LEE: added as check
                                                ReDim Preserve arrGG(intGG)
                                            End If
                                            arrGG(intGG) = Count6
                                        End If
                                    End If

                                Next Count6

                            Next Count5

                            'continue evaluating
                            'arrGG has group #'s that have same number of levels as dv3
                            'determine if any of these group's concentration levels are different

                            Dim arrGGG()
                            ReDim arrGGG(intGG)
                            Dim intGGG As Short = 0
                            Dim strIS1 As String
                            Dim strIS2 As String

                            boolHit1 = False
                            boolHit = False
                            For Count6 = 1 To intGG

                                int1 = arrGG(Count6)
                                'strF6 = "INTGROUP = " & int1
                                strF6 = "INTGROUP = " & int1 & " AND ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                                boolHit = False
                                rows5 = tblCalStdGroupsAll.Select(strF6, "CONCENTRATION ASC")

                                'compare concentrations and levelnumber
                                'must be true at all levels
                                'examine all existing arrGG's
                                If rows5.Length = 0 Then
                                Else

                                    boolHit = False
                                    For Count7 = 0 To rows5.Length - 1


                                        If rows2.Length = 0 Then

                                            '20171219 LEE

                                            c1 = -1
                                            l1 = -1
                                            strIS1 = ""
                                        Else
                                            c1 = NZ(rows2(Count7).Item("CONCENTRATION"), -1)
                                            l1 = NZ(rows2(Count7).Item("LEVELNUMBER"), -1)
                                            strIS1 = NZ(rows2(Count7).Item("INTERNALSTANDARD"), "")
                                        End If

                                        c2 = NZ(rows5(Count7).Item("CONCENTRATION"), -1)
                                        l2 = NZ(rows5(Count7).Item("LEVELNUMBER"), -1)
                                        strIS2 = NZ(rows5(Count7).Item("INTSTD"), "NA")

                                        If boolDiffIS Then

                                            '20180316 LEE:
                                            'Hmm. Should not be looking at levels
                                            'Levels can be different but nomconc the same

                                            'If c1 = c2 And l1 = l2 And StrComp(strIS1, strIS2, CompareMethod.Text) = 0 Then
                                            If c1 = c2 And StrComp(strIS1, strIS2, CompareMethod.Text) = 0 Then
                                            Else
                                                boolHit = True
                                                Exit For
                                            End If
                                        Else
                                            If c1 = c2 And l1 = l2 Then
                                            Else
                                                boolHit = True
                                                Exit For
                                            End If
                                        End If


                                    Next

                                    If boolHit Or rows2.Length = 0 Then
                                        If rows2.Length = 0 Then

                                            '20171219 LEE

                                            boolHit = False
                                        End If
                                        var1 = var1
                                    Else
                                        'found a match
                                        'must check to see if RegressionStatus and PSAE (runtypeid = 3) is different
                                        var1 = rows5(0).Item("RUNANALYTEREGRESSIONSTATUS")
                                        var2 = rows5(0).Item("RUNID")
                                        var3 = rows5(0).Item("RUNTYPEID")
                                        'need to record this group's intCalSet
                                        intCalSetR = rows5(0).Item("INTCALSET")
                                        'If var1 <> 3 And intRARS = 3 And var3 = 3 Then
                                        If intRARS = 3 And (var3 = 3 Or var1 <> 3) And rows2.Length <> 0 Then
                                            'must replace with this one

                                            Try
                                                var1 = var1 'debug
                                                For Count8 = 0 To rows5.Length - 1

                                                    rows5(Count8).BeginEdit()

                                                    rows5(Count8).Item("ANALYTEINDEX") = intAnalyteIndex ' 
                                                    rows5(Count8).Item("MASTERASSAYID") = intMasterAssayID
                                                    rows5(Count8).Item("RUNID") = intRunID
                                                    rows5(Count8).Item("RUNTYPEID") = intRunTypeID
                                                    rows5(Count8).Item("ASSAYID") = intAssayID
                                                    rows5(Count8).Item("RUNANALYTEREGRESSIONSTATUS") = intRARS
                                                    rows5(Count8).Item("RUNTYPE") = strRunType
                                                    rows5(Count8).Item("LLOQ") = numLLOQ
                                                    rows5(Count8).Item("ULOQ") = numULOQ
                                                    rows5(Count8).Item("CONCENTRATIONUNITS") = strConcUnits '
                                                    rows5(Count8).Item("INTSTD") = strIntStdName ' 
                                                    rows5(Count8).Item("RUNDATE") = dtRunDate ' 


                                                    'intAnalyteIndex = dvUAssID(Count2).Item("ANALYTEINDEX")
                                                    'intMasterAssayID = dvUAssID(Count2).Item("MASTERASSAYID") 'probably don't need this, but record anyway
                                                    'intRunID = dvUAssID(Count2).Item("RUNID")
                                                    'intRunTypeID = dvUAssID(Count2).Item("RUNTYPEID")
                                                    'intAssayID = dvUAssID(Count2).Item("ASSAYID")
                                                    'intRARS = dvUAssID(Count2).Item("RUNANALYTEREGRESSIONSTATUS")
                                                    'numLLOQ = rowAQL(0).Item("NM")
                                                    'numULOQ = rowAQL(0).Item("VEC")
                                                    'strConcUnits = rowAQL(0).Item("CONCENTRATIONUNITS")
                                                    'strIntStdName = NZ(rowIntStd(0).Item("INTERNALSTANDARD"),"NA")
                                                    'dtRunDate = rows5(0).Item("RUNSTARTDATE")

                                                    rows5(Count8).EndEdit()

                                                Next
                                            Catch ex As Exception
                                                var1 = ex.Message
                                                var1 = var1
                                            End Try

                                        End If
                                        Exit For
                                    End If

                                End If

                            Next

                            If boolHit = False Then 'a group is equal, so ignore
                                intGroupA = int1
                            Else

                                'make new group
                                intGroup = intGroup + 1
                                intGroupA = intGroup
                                intCalSet = intCalSet + 1
                                gCalSet = intCalSet

                                If rows2.Length = 0 Then

                                    '20171219 LEE

                                    intLevelNumber = -1
                                    numNomConc = -1
                                    numAFP = 15

                                    Dim nr As DataRow = tblCalStdGroupsAll.NewRow
                                    nr.BeginEdit()
                                    nr.Item("ANALYTEDESCRIPTION") = strAnalyteName
                                    nr.Item("INTSTD") = strIntStdName
                                    nr.Item("ANALYTEID") = intAnalyteID
                                    nr.Item("ANALYTEINDEX") = intAnalyteIndex
                                    nr.Item("ASSAYID") = intAssayID
                                    nr.Item("MASTERASSAYID") = intMasterAssayID
                                    nr.Item("LEVELNUMBER") = intLevelNumber
                                    nr.Item("CONCENTRATION") = numNomConc

                                    nr.Item("ANALYTEFLAGPERCENT") = numAFP

                                    nr.Item("INTGROUP") = intGroup
                                    nr.Item("RUNTYPEID") = intRunTypeID
                                    nr.Item("RUNANALYTEREGRESSIONSTATUS") = intRARS

                                    nr.Item("RUNTYPE") = strRunType

                                    nr.Item("LLOQ") = numLLOQ
                                    nr.Item("ULOQ") = numULOQ
                                    nr.Item("CONCENTRATIONUNITS") = strConcUnits

                                    nr.Item("RUNID") = intRunID
                                    nr.Item("RUNDATE") = dtRunDate
                                    nr.Item("MATRIX") = strMatrix
                                    nr.Item("INTCALSET") = intCalSet

                                    nr.EndEdit()

                                    Try
                                        tblCalStdGroupsAll.Rows.Add(nr)
                                    Catch ex As Exception
                                        var1 = ex.Message
                                    End Try
                                Else
                                    For Count3 = 0 To rows2.Length - 1

                                        intLevelNumber = NZ(rows2(Count3).Item("LEVELNUMBER"), -1)
                                        numNomConc = NZ(rows2(Count3).Item("CONCENTRATION"), -1)
                                        numAFP = NZ(rows2(Count3).Item("ANALYTEFLAGPERCENT"), NZ(rows2(Count3).Item("FLAGPERCENT"), 15))

                                        Dim nr As DataRow = tblCalStdGroupsAll.NewRow
                                        nr.BeginEdit()
                                        nr.Item("ANALYTEDESCRIPTION") = strAnalyteName
                                        nr.Item("INTSTD") = strIntStdName
                                        nr.Item("ANALYTEID") = intAnalyteID
                                        nr.Item("ANALYTEINDEX") = intAnalyteIndex
                                        nr.Item("ASSAYID") = intAssayID
                                        nr.Item("MASTERASSAYID") = intMasterAssayID
                                        nr.Item("LEVELNUMBER") = intLevelNumber
                                        nr.Item("CONCENTRATION") = numNomConc

                                        nr.Item("ANALYTEFLAGPERCENT") = numAFP

                                        nr.Item("INTGROUP") = intGroup
                                        nr.Item("RUNTYPEID") = intRunTypeID
                                        nr.Item("RUNANALYTEREGRESSIONSTATUS") = intRARS

                                        nr.Item("RUNTYPE") = strRunType

                                        nr.Item("LLOQ") = numLLOQ
                                        nr.Item("ULOQ") = numULOQ
                                        nr.Item("CONCENTRATIONUNITS") = strConcUnits

                                        nr.Item("RUNID") = intRunID
                                        nr.Item("RUNDATE") = dtRunDate
                                        nr.Item("MATRIX") = strMatrix
                                        nr.Item("INTCALSET") = intCalSet

                                        nr.EndEdit()

                                        Try
                                            tblCalStdGroupsAll.Rows.Add(nr)
                                        Catch ex As Exception
                                            var1 = ex.Message
                                        End Try

                                    Next Count3
                                End If



                                intCalSetR = intCalSet

                                var1 = var1 'debug
                                int1 = tblCalStdGroupsAll.Rows.Count
                                int1 = int1

                            End If

                        End If

                        'enter record in tblCalStdGroupAssayIDsAll
                        'ensure record doesn't already exist
                        'strF7 = "ASSAYID = " & intAssayID & " AND ANALYTEID = " & intAnalyteID & " AND ANALYTEINDEX = " & intAnalyteIndex
                        strF7 = "ASSAYID = " & intAssayID & " AND ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "'"
                        rows6 = tblCalStdGroupAssayIDsAll.Select(strF7)

                        If rows6.Length = 0 Then

                            'need to find group
                            If intCalSetR = 0 Then
                                var1 = var1
                            End If
                            strF = "ANALYTEID = " & intAnalyteID & " AND MATRIX = '" & strMatrix & "' AND INTCALSET = " & intCalSetR
                            Dim rowsXX() As DataRow = tblCalStdGroupsAll.Select(strF)
                            'Try
                            '    rowsXX = tblCalStdGroupsAll.Select(strF)
                            'Catch ex As Exception
                            '    var1 = ex.Message
                            'End Try

                            If rowsXX.Length = 0 Then
                                intGroupD = intGroup
                            Else
                                intGroupD = rowsXX(0).Item("INTGROUP")
                            End If


                            'record tblCalStdGroupAssayIDsAll
                            Dim nr1 As DataRow = tblCalStdGroupAssayIDsAll.NewRow
                            nr1.BeginEdit()

                            nr1.Item("ANALYTEDESCRIPTION") = strAnalyteName
                            nr1.Item("INTSTD") = strIntStdName
                            nr1.Item("ANALYTEID") = intAnalyteID
                            nr1.Item("ANALYTEINDEX") = intAnalyteIndex
                            nr1.Item("ASSAYID") = intAssayID
                            nr1.Item("INTGROUP") = intGroupD ' intGroupA
                            nr1.Item("RUNTYPEID") = intRunTypeID
                            nr1.Item("RUNANALYTEREGRESSIONSTATUS") = intRARS

                            nr1.Item("RUNTYPE") = strRunType

                            nr1.Item("MASTERASSAYID") = intMasterAssayID

                            nr1.Item("LLOQ") = numLLOQ
                            nr1.Item("ULOQ") = numULOQ
                            nr1.Item("CONCENTRATIONUNITS") = strConcUnits

                            nr1.Item("RUNID") = intRunID
                            nr1.Item("RUNDATE") = dtRunDate
                            nr1.Item("MATRIX") = strMatrix
                            nr1.Item("INTCALSET") = intCalSetR

                            nr1.EndEdit()
                            Try
                                tblCalStdGroupAssayIDsAll.Rows.Add(nr1)
                            Catch ex As Exception
                                var1 = ex.Message
                            End Try

                        End If
NextCount2:


                    Next Count2
                Catch ex As Exception
                    var1 = var1
                End Try

                

                int1 = tblCalStdGroupsAll.Rows.Count
                int1 = int1

            Next CountID

        Next CountM

        'now create tblCalStdGroupsAcc and   tblCalStdGroupAssayIDsAcc
        '(RUNANALYTEREGRESSIONSTATUS=3) AND (RUNTYPEID<>3) ' RUNANALYTEREGRESSIONSTATUS=3=accepted  RunTypeID = 3 = PSAE
        Try
            strF = "RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <>3 "
            Dim dvAA As DataView = New DataView(tblCalStdGroupsAll, strF, "", DataViewRowState.CurrentRows)
            tblCalStdGroupsAcc = dvAA.ToTable

        Catch ex As Exception
            var1 = ex.Message
        End Try

        Try
            strF = "RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <>3 "
            Dim dvAB As DataView = New DataView(tblCalStdGroupAssayIDsAll, strF, "", DataViewRowState.CurrentRows)
            tblCalStdGroupAssayIDsAcc = dvAB.ToTable
            var1 = var1 'debug
        Catch ex As Exception
            var1 = ex.Message
        End Try

        'if cal std, create tblanalytegroups
        If boolStd Then

            'get unique assayid's from this table
            'set to dataview
            '20160217 LEE: gSortAnalytes is public that may be allowed to be modified by user in the future
            'for now, default = 'Matrix'
            strS = ReturnSort(True)
            gSortAnalyteString = strS 'use this in Reassay, Repeat, and Sample Conc tables

            Dim dv1a As DataView = New DataView(tblCalStdGroupsAll, "", strS, DataViewRowState.CurrentRows)
            Try
                tblAnalyteGroupsTemp = dv1a.ToTable("a", True, "ANALYTEDESCRIPTION", "ANALYTEID", "INTSTD", "INTGROUP", "ANALYTEDESCRIPTION_C", "MATRIX", "INTCALSET", "CALIBRSET")
            Catch ex As Exception
                var1 = ex.Message
            End Try

            'now add a column
            If tblAnalyteGroupsTemp.Columns.Contains("INTORDER") Then
            Else
                Dim col10 As New DataColumn
                str1 = "INTORDER"
                col10.DataType = System.Type.GetType("System.Int16")
                col10.ColumnName = str1
                col10.Caption = str1 'strAnal
                tblAnalyteGroupsTemp.Columns.Add(col10)

                'now add values
                For Count1 = 0 To tblAnalyteGroupsTemp.Rows.Count - 1
                    'debug
                    var1 = tblAnalyteGroupsTemp.Rows(Count1).Item("INTGROUP")
                    tblAnalyteGroupsTemp.Rows(Count1).BeginEdit()
                    tblAnalyteGroupsTemp.Rows(Count1).Item("INTORDER") = Count1 + 1
                    tblAnalyteGroupsTemp.Rows(Count1).EndEdit()
                Next
            End If

            If tblAnalyteGroupsTemp.Columns.Contains("CHARUSERANALYTE") Then
            Else
                Dim col6f As New DataColumn
                str1 = "CHARUSERANALYTE"
                col6f.DataType = System.Type.GetType("System.String")
                col6f.ColumnName = str1
                col6f.Caption = str1 'strAnal
                tblAnalyteGroupsTemp.Columns.Add(col6f)

            End If

            If tblAnalyteGroupsTemp.Columns.Contains("CHARUSERANALYTE") Then
            Else
                Dim col6g As New DataColumn
                str1 = "CHARUSERIS"
                col6g.DataType = System.Type.GetType("System.String")
                col6g.ColumnName = str1
                col6g.Caption = str1 'strAnal
                tblAnalyteGroupsTemp.Columns.Add(col6g)

            End If


            'fill CALIBRSET
            'tblCalStdGroupAssayIDsAll
            Dim rowSGA() As DataRow
            For Count1 = 1 To tblAnalyteGroupsTemp.Rows.Count
                strF = "INTGROUP = " & tblAnalyteGroupsTemp.Rows(Count1 - 1).Item("INTGROUP")
                strS = "CONCENTRATION ASC"
                rowSGA = tblCalStdGroupsAll.Select(strF, strS)

                'now make unique Concentration and LevelNumber
                Dim tblU1 As DataTable = rowSGA.CopyToDataTable
                Dim dvU1 As DataView = New DataView(tblU1, "", "", DataViewRowState.CurrentRows)
                'Hmm. Should not be looking at levels
                'Levels can be different but nomconc the same
                'Dim tblU2 As DataTable = dvU1.ToTable("b", True, "CONCENTRATION", "LEVELNUMBER")
                Dim tblU2 As DataTable = dvU1.ToTable("b", True, "CONCENTRATION")

                For Count2 = 1 To tblU2.Rows.Count
                    var1 = tblU2.Rows(Count2 - 1).Item("CONCENTRATION")

                    '20180316
                    'Hmm. Should not be looking at levels
                    'Levels can be different but nomconc the same

                    'var3 = tblU2.Rows(Count2 - 1).Item("LEVELNUMBER")
                    'need to find level number in dv1a
                    'dv1a.RowFilter = strF & " AND CONCENTRATION = " & var1 & " AND LEVELNUMBER = " & var3
                    dv1a.RowFilter = strF & " AND CONCENTRATION = " & var1
                    var2 = dv1a.Count 'debug
                    var2 = dv1a(0).Item("LEVELNUMBER")
                    If IsDBNull(var1) Then
                        var1 = var1
                    End If
                    'str1 = rowSGA(Count2 - 1).Item("CONCENTRATION").ToString
                    '20181130 LEE:
                    str1 = tblU2.Rows(Count2 - 1).Item("CONCENTRATION").ToString
                    str3 = var2.ToString
                    If Count2 = 1 Then
                        str2 = str1 & "(" & str3 & ")"
                    Else
                        str2 = str2 & ", " & str1 & "(" & str3 & ")"
                    End If
                Next

                'For Count2 = 1 To rowSGA.Length
                '    var1 = rowSGA(Count2 - 1).Item("CONCENTRATION")
                '    'need to find level number in dv1a
                '    dv1a.RowFilter = strF & " AND CONCENTRATION = " & var1
                '    var2 = dv1a.Count 'debug
                '    var2 = dv1a(0).Item("LEVELNUMBER")
                '    If IsDBNull(var1) Then
                '        var1 = var1
                '    End If
                '    str1 = rowSGA(Count2 - 1).Item("CONCENTRATION").ToString
                '    str3 = var2.ToString
                '    If Count2 = 1 Then
                '        str2 = str1 & "(" & str3 & ")"
                '    Else
                '        str2 = str2 & ", " & str1 & "(" & str3 & ")"
                '    End If
                'Next
                tblAnalyteGroupsTemp.Rows(Count1 - 1).BeginEdit()
                tblAnalyteGroupsTemp.Rows(Count1 - 1).Item("CALIBRSET") = str2
                tblAnalyteGroupsTemp.Rows(Count1 - 1).EndEdit()
            Next

            Dim intAGs As Int16 = tblAnalyteGroupsTemp.Rows.Count

            'now check for duplicates
            Dim dv1b As DataView = New DataView(tblAnalyteGroupsTemp, "", "", DataViewRowState.CurrentRows)
            Dim tblAG As System.Data.DataTable = dv1b.ToTable("b", True, "ANALYTEDESCRIPTION")
            Dim intAG As Int16 = tblAG.Rows.Count
            dv1b.Sort = "ANALYTEDESCRIPTION ASC"

            Dim boolCalSet As Boolean = False
            Dim boolMatrixSet As Boolean = False
            Dim strEntry As String

            If intNumMatrixHere > 1 Then
                boolMatrixSet = True
            End If

            If intAGs = intAG Then 'no duplicates
                For Count1 = 0 To intAG - 1
                    tblAnalyteGroupsTemp.Rows(Count1).BeginEdit()
                    tblAnalyteGroupsTemp.Rows(Count1).Item("ANALYTEDESCRIPTION_C") = tblAnalyteGroupsTemp.Rows(Count1).Item("ANALYTEDESCRIPTION")
                    tblAnalyteGroupsTemp.Rows(Count1).EndEdit()
                Next
            Else
                For Count3 = 1 To intNumMatrixHere

                    strMatrix = tblSP.Rows(Count3 - 1).Item("SAMPLETYPEID")

                    For Count5 = 1 To numAnalIDs

                        intAnalyteID = tblAnalIDs.Rows(Count5 - 1).Item("ANALYTEID")
                        strAnalyteName = tblAnalIDs.Rows(Count5 - 1).Item("ANALYTEDESCRIPTION")

                        strF = "ANALYTEDESCRIPTION = '" & CleanText(strAnalyteName) & "' AND MATRIX = '" & strMatrix & "'"
                        Dim dvCSet As DataView = New DataView(tblAnalyteGroupsTemp, strF, "", DataViewRowState.CurrentRows)
                        Dim tblCSet As DataTable = dvCSet.ToTable("cs", True, "INTCALSET")

                        intCalSet = tblCSet.Rows.Count

                        If intCalSet > 1 Then
                            boolCalSet = True
                        Else
                            boolCalSet = False
                        End If

                        For Count2 = 1 To intCalSet

                            'For Count1 = 0 To intAG - 1

                            str1 = strAnalyteName ' tblAG.Rows(Count1).Item("ANALYTEDESCRIPTION")
                            strF = "ANALYTEDESCRIPTION = '" & CleanText(str1) & "' AND MATRIX = '" & strMatrix & "' AND INTCALSET = " & Count2
                            Dim rows10() As DataRow = tblAnalyteGroupsTemp.Select(strF, "ANALYTEDESCRIPTION_C ASC, INTGROUP ASC")

                            If boolMatrixSet And boolCalSet Then
                                str4 = str1 & "_M" & Count3 & "_C" & Count2
                                str4 = str1 & " " & strMatrix & "_C" & Count2
                            ElseIf boolMatrixSet Then
                                str4 = str1 & " " & strMatrix
                            ElseIf boolCalSet Then
                                str4 = str1 & "_C" & Count2
                            Else
                                str4 = str1
                            End If

                            For Count4 = 0 To rows10.Length - 1
                                str2 = str4
                                rows10(Count4).BeginEdit()
                                rows10(Count4).Item("ANALYTEDESCRIPTION_C") = str2
                                rows10(Count4).EndEdit()
                            Next Count4

                            'Next Count1


                        Next Count2

                    Next Count5

                Next Count3


                ''now fill ANALYTEDESCRIPTION_C in tblCalStdGroupsAll
                'For Count1 = 0 To intAGs - 1
                '    int1 = tblAnalyteGroups.Rows(Count1).Item("INTGROUP")
                '    str2 = tblAnalyteGroups.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
                '    strF = "INTGROUP = " & int1
                '    Dim rows11() As DataRow = tblCalStdGroupsAll.Select(strF)
                '    For Count2 = 0 To rows11.Length - 1
                '        rows11(Count2).BeginEdit()
                '        rows11(Count2).Item("ANALYTEDESCRIPTION_C") = str2
                '        rows11(Count2).EndEdit()
                '    Next
                'Next

            End If

            var1 = intAG 'debug

        Else

        End If

        If gCalSet = 1 Then
            frmH.cmdShowGroups.Visible = False
        Else
            frmH.cmdShowGroups.Visible = True
        End If

        var1 = tblAnalyteGroupsTemp.Rows.Count

        'now fill ANALYTEDESCRIPTION_C in tblCalStdGroupsAll

        For Count1 = 0 To tblAnalyteGroupsTemp.Rows.Count - 1
            int1 = tblAnalyteGroupsTemp.Rows(Count1).Item("INTGROUP")
            str2 = tblAnalyteGroupsTemp.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
            str4 = tblAnalyteGroupsTemp.Rows(Count1).Item("ANALYTEDESCRIPTION")
            'int2 = tblAnalyteGroups.Rows(Count1).Item("ANALYTEID")
            'int3 = tblAnalyteGroups.Rows(Count1).Item("ANALYTEINDEX")
            strF = "INTGROUP = " & int1

            Dim rows11() As DataRow = tblCalStdGroupsAll.Select(strF)
            For Count2 = 0 To rows11.Length - 1
                str3 = rows11(Count2).Item("ANALYTEDESCRIPTION") 'check
                rows11(Count2).BeginEdit()
                rows11(Count2).Item("ANALYTEDESCRIPTION_C") = str2
                rows11(Count2).EndEdit()
            Next
        Next

        'now fill ANALYTEDESCRIPTION_C in tblCalStdGroupAssayIDsAll
        'must fill from tblCalStdGroupsAll
        For Count1 = 0 To tblAnalyteGroupsTemp.Rows.Count - 1
            int1 = tblAnalyteGroupsTemp.Rows(Count1).Item("INTGROUP")
            str2 = tblAnalyteGroupsTemp.Rows(Count1).Item("ANALYTEDESCRIPTION_C")
            str4 = tblAnalyteGroupsTemp.Rows(Count1).Item("ANALYTEDESCRIPTION")
            strF = "INTGROUP = " & int1
            Dim rows11() As DataRow = tblCalStdGroupAssayIDsAll.Select(strF)
            For Count2 = 0 To rows11.Length - 1
                str3 = rows11(Count2).Item("ANALYTEDESCRIPTION") 'check
                rows11(Count2).BeginEdit()
                rows11(Count2).Item("ANALYTEDESCRIPTION_C") = str2
                rows11(Count2).EndEdit()
            Next
        Next

        'now create tblCalStdGroupsAcc and   tblCalStdGroupAssayIDsAcc
        '(RUNANALYTEREGRESSIONSTATUS=3) AND (RUNTYPEID<>3) ' RUNANALYTEREGRESSIONSTATUS=3=accepted  RunTypeID = 3 = PSAE
        Try
            strF = "RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <>3 "
            Dim dvAA As DataView = New DataView(tblCalStdGroupsAll, strF, "", DataViewRowState.CurrentRows)
            tblCalStdGroupsAcc = dvAA.ToTable
            var1 = tblCalStdGroupsAcc.Rows.Count 'debug
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Try
            Dim dvAB As DataView = New DataView(tblCalStdGroupAssayIDsAll, strF, "", DataViewRowState.CurrentRows)
            tblCalStdGroupAssayIDsAcc = dvAB.ToTable
            var1 = tblCalStdGroupAssayIDsAcc.Rows.Count 'debug
        Catch ex As Exception
            var1 = ex.Message
        End Try

        'now reset tblAnalyteGroups to have only Accepted Analytical Runs
        strF = "RUNANALYTEREGRESSIONSTATUS = 3"
        Dim dv1c As DataView = New DataView(tblCalStdGroupsAll, strF, "ANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION_C ASC, INTGROUP ASC", DataViewRowState.CurrentRows)
        tblAnalyteGroupsAcc = dv1c.ToTable("b", True, "ANALYTEDESCRIPTION", "ANALYTEID", "INTSTD", "INTGROUP", "ANALYTEDESCRIPTION_C", "MATRIX", "INTCALSET", "CALIBRSET")

        tblAnalyteGroupsTemp.AcceptChanges()
        tblAnalyteGroupsAcc.AcceptChanges()

        'do this temporarily
        tblAnalyteGroups = tblAnalyteGroupsTemp
        var1 = tblAnalyteGroups.Rows.Count 'debug
        var1 = var1

        'now must sinc with StudyDoc tblAnalyteGroups

        'assign dgvGroups
        Call dgvSetGroups(frmH.dgvGroups)


end1:

        'if you wish to log data in console, comment out next line
        GoTo end2

        Dim dtbl As DataTable
        Dim dv As DataView
        Dim rows() As System.Data.DataRow
        Try
            For Count1 = 1 To 1
                Select Case Count1
                    Case Is = 1
                        str1 = "tblAnalyteGroups"
                        dv = New DataView(tblAnalyteGroups, "INTGROUP > 0", "INTGROUP ASC", DataViewRowState.CurrentRows)
                        dtbl = dv.ToTable
                    Case Is = 2
                        str1 = "tblCalStdGroupsAll"
                        dtbl = tblCalStdGroupsAll
                    Case Is = 3
                        str1 = "tblCalStdGroupsAcc"
                        dtbl = tblCalStdGroupsAcc
                    Case Is = 4
                        str1 = "tblCalStdGroupAssayIDsAll"
                        dv = New DataView(tblCalStdGroupAssayIDsAll, "INTGROUP > 0", "INTGROUP ASC, RUNID ASC", DataViewRowState.CurrentRows)
                        dtbl = dv.ToTable
                    Case Is = 5
                        str1 = "tblCalStdGroupAssayIDsAcc"
                        dv = New DataView(tblCalStdGroupAssayIDsAcc, "INTGROUP > 0", "INTGROUP ASC, RUNID ASC", DataViewRowState.CurrentRows)
                        dtbl = dv.ToTable
                End Select

                Console.WriteLine(" ")
                Console.WriteLine("Begin " & str1)
                'do headings
                str3 = ""
                For Count2 = 0 To dtbl.Columns.Count - 1
                    str4 = dtbl.Columns(Count2).ColumnName
                    If Count2 = 0 Then
                        str3 = str4
                    Else
                        str3 = str3 & ChrW(9) & str4
                    End If
                Next
                Console.WriteLine(str3)

                'do data
                str5 = ""
                For Count2 = 0 To dtbl.Rows.Count - 1
                    str3 = ""
                    str5 = ""
                    For Count3 = 0 To dtbl.Columns.Count - 1
                        str4 = NZ(dtbl.Rows(Count2).Item(Count3), " ")
                        If Count3 = 0 Then
                            str3 = str4
                        Else
                            str3 = str3 & ChrW(9) & str4
                        End If
                    Next
                    If Count2 = 0 Then
                        str5 = str3
                    Else
                        str5 = str5 & ChrW(10) & str3
                    End If

                    Console.WriteLine(str3)
                Next

                var1 = var1 'debug

            Next
            Console.WriteLine(" ")
        Catch ex As Exception
            var1 = ex.Message
        End Try

end2:

        da.Dispose()

    End Sub



End Module
