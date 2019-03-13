Option Compare Text

Public Class frmUpdateCheck

    Public boolBrandNew As Boolean
    Public boolAccess As Boolean
    Public boolOra As Boolean
    Public boolSQLSer As Boolean

    Private Sub frmUpdateCheck_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim strM As String

        If boolBrandNew Then
            Me.pan1.Visible = True
            Me.cmdYes.Visible = True
            Me.cmdNo.Visible = True
            Me.cmdNo.Enabled = False

            strM = Me.lbl1.Text
            strM = strM & ChrW(10) & ChrW(10) & "Please click on 'Yes' to continue."
            Me.lbl1.Text = strM

        Else
            Me.pan1.Visible = True
            'Me.cmdYes.Visible = True
            'Me.cmdNo.Visible = True
        End If

end1:

    End Sub

    Sub DoButton(ByVal strB As String)
        'DoButton: Performs actions to update StudyDoc Database if user answers "Yes"

        If StrComp(strB, "Yes", CompareMethod.Text) = 0 Then

            Call CheckTables(boolAccess, boolOra, boolSQLSer)

            boolUpdateCheckBad = False

            boolBad = False
            If boolNeedsUD Then
                If boolBad Then

                Else
                    strUpdateMsg = "StudyDoc database updated successfully."
                    strUpdateMsg = strUpdateMsg & ChrW(10) & ChrW(10) & "StudyDoc will close in order to finalize database changes."
                    strUpdateMsg = strUpdateMsg & ChrW(10) & ChrW(10) & "Please re-start StudyDoc."

                    MsgBox(strUpdateMsg, MsgBoxStyle.Information, "Finished...")

                    End


                    boolBad = False
                End If

            Else
                boolBad = False

            End If

            boolUpdateCheckBad = boolBad
        Else

            End

        End If

end1:
        Me.Close()


    End Sub

    Private Sub cmdYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYes.Click

        Cursor.Current = Cursors.WaitCursor

        Call DoButton("Yes")
        boolQuitUpdate = False

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdNo_Click(sender As Object, e As EventArgs) Handles cmdNo.Click

        Call DoButton("No")
        boolQuitUpdate = True

    End Sub

    Sub CheckTables(ByVal boolAccess As Boolean, ByVal boolOra As Boolean, ByVal boolSQLSer As Boolean)
        'CheckTables:  Check StudyDoc Database table versions, and update the tables if they are out of date.

        Dim constrB As String
        Dim strPath As String
        Dim strPathDB As String
        Dim strPathDBmdb As String
        Dim strMDB As String
        Dim rsS As New ADODB.Recordset
        Dim rsD As New ADODB.Recordset
        Dim rsMaxID As New ADODB.Recordset
        Dim rsTables As New ADODB.Recordset

        Dim strT As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim boolProb As Boolean = True
        Dim boolCont As Boolean = True
        Dim boolExists As Boolean = False
        Dim strM As String
        Dim intRec As Short
        Dim Count1 As Short
        Dim Count2 As Int16
        Dim boolExact As Boolean
        Dim boolExactValues As Boolean
        Dim boolMultPK As Boolean
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim boolSpecial As Boolean
        Dim strPathMDBDest As String
        Dim var1, var2, var3
        Dim intTNum As Short = 0

        Me.pgOverall.Maximum = 3
        Me.pgOverall.Value = 0
        Me.pgOverall.Refresh()

        '20190131 LEE:
        'initialize strInstall
        intInstall = 0
        ReDim strInstall(2000000)

        '20150822 Larry: This code isn't needed anymore
        'From Here:
        'get path of current executable
        strMDB = "GuWu_01.mdb"
        strMDB = "StudyDoc_01.mdb"
        strPath = System.Windows.Forms.Application.ExecutablePath
        'strip off file name
        For Count1 = Len(strPath) To 1 Step -1
            str1 = Mid(strPath, Count1, 1)
            If StrComp(str1, "\", CompareMethod.Text) = 0 Then
                Exit For
            End If
        Next
        strPath = Mid(strPath, 1, Count1) 'has a backslash at the end
        strPathDB = strPath & "DatabaseInstallation\"
        strPathDBmdb = strPathDB & strMDB
        'create constrB
        'for testing
        str1 = "C:\Program Files"
        If InStr(1, strPathDBmdb, str1, CompareMethod.Text) > 0 Then
        Else
            strPathDBmdb = "C:\Labintegrity\StudyDoc\MDBDatabase\GuWu_01.mdb"
            strPathDBmdb = "C:\LabIntegrity\StudyDoc\MDBDatabase\StudyDoc_01.mdb"
        End If

        If boolDemo Then
            strPathDBmdb = "C:\Labintegrity\StudyDoc\MDBDatabase\GuWu_01.mdb"
            strPathDBmdb = "C:\LabIntegrity\StudyDoc\MDBDatabase\StudyDoc_01.mdb"
        End If

        'force this next line until dsn gets configured
        strPathDBmdb = "C:\Labintegrity\StudyDoc\MDBDatabase\GuWu_01.mdb"
        strPathDBmdb = "C:\LabIntegrity\StudyDoc\MDBDatabase\StudyDoc_01.mdb"
        'NO! Database may be located on network!!!!
        'constrB should be constrini

        constrB = constrIni ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPathDBmdb & ";"
        If boolGuWuAccess Then
            constrB = constrIni ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPathDBmdb & ";"
        ElseIf boolGuWuSQLServer Then
            constrB = constrIni '"Provider=SQLOLEDB;" & constrIni ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPathDBmdb & ";"
        ElseIf boolGuWuOracle Then
            constrB = constrIniGuWuODBC
        End If


        If boolAccess Then
            'first backup Access database
            'retrieve path of Access database from constrini
            'E.g:  Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\LabIntegrity\StudyDoc\MDBDatabase\StudyDoc_01.mdb;
            int3 = InStr(1, constrIni, ";", CompareMethod.Text)
            int1 = InStr(int3, constrIni, "=", CompareMethod.Text)
            int2 = Len(constrIni)
            Dim strPathMDB As String
            Dim strPathMDBBU As String
            strPathMDB = Mid(constrIni, int1 + 1, int2 - int1 - 1)
            var1 = "_" & strVerOld & "_" & Format(Now, "yyyyMMdd_HHmmss") & ".mdb"
            strPathMDBBU = Replace(strPathMDB, ".mdb", var1, 1, -1, CompareMethod.Text)

            str1 = Me.lbl1.Text
            str2 = str1 & ChrW(10) & ChrW(10) & "Backing Up StudyDoc database to:" & ChrW(10) & strPathMDBBU
            Me.lbl1.Text = str2
            Me.lbl1.Refresh()

            Try
                System.IO.File.Copy(strPathMDB, strPathMDBBU)
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try


            Pause(0.5)
            'check to ensure copy completed
            If System.IO.File.Exists(strPathMDBBU) Then
                str2 = str1 & ChrW(10) & ChrW(10) & "Successful Backing Up StudyDoc database to:" & ChrW(10) & strPathMDBBU
                Me.lbl1.Text = str2
                Me.lbl1.Refresh()
                Pause(0.5)
            Else
                strM = "There was a problem backing up the existing StudyDoc database to:" & ChrW(10) & ChrW(10) & strPathMDBBU & ChrW(10) & ChrW(10)
                strM = strM & "Click 'Yes' if the StudyDoc database has been backed up manually." & ChrW(10) & ChrW(10)
                strM = strM & "Click 'No' to stop this update action."
                Dim intR As Short = MsgBox(strM, vbYesNo, "Continue?")
                If intR = 6 Then 'continue
                Else
                    End
                End If
            End If

            'str1 = Me.lbl1.Text
            'str2 = str1 & ChrW(10) & ChrW(10) & "Completed Backing Up StudyDoc database to:" & ChrW(10) & ChrW(10) & strPathMDBBU
            Me.lbl1.Text = str1
            Me.lbl1.Refresh()

        End If

        'open conb

        'constrini: DATA SOURCE=GUBBSLAP07;INITIAL CATALOG=STUDYDOC_01SQL;UID=Gubbs;PWD=Gwoman1;
        'constrB: Provider=SQLOLEDB;DATA SOURCE=GUBBSLAP07;INITIAL CATALOG=STUDYDOC_01SQL;UID=Gubbs;PWD=Gwoman1;

        '20190205 LEE:
        'need to use datatable, Frontage cannot updatebatch with adodb
        Dim tblVer As DataTable
        Try
            tblVer = tblVersion
        Catch ex As Exception
            var1 = var1
        End Try
        var1 = tblVer.Rows.Count

        conB.Open(constrB)


        Dim myConnection As OleDb.OleDbConnection
        Dim myCommand As New OleDb.OleDbCommand

        Try
            ''console.writeline("constrini: " & constrIni)
            ''console.writeline("constrB: " & constrB)
            If boolAccess Then
                myConnection = New OleDb.OleDbConnection(constrIni)
            Else
                myConnection = New OleDb.OleDbConnection(constrB)
            End If

            myConnection.Open()
            myCommand.Connection = myConnection
            myCommand.CommandType = CommandType.Text
        Catch ex As Exception
            var1 = ex.Message
        End Try


        If boolAccess Then

            'get path from conA.connectionstring
            str1 = conA.ConnectionString.ToString
            str2 = "Data Source="
            int1 = InStr(1, str1, str2, CompareMethod.Text)
            int2 = InStr(int1 + Len(str2) + 1, str1, ";", CompareMethod.Text)
            str3 = Mid(str1, int1 + Len(str2), int2 - (int1 + Len(str2)))
            'int2 = InStr(int1 + 1, str1, ";", CompareMethod.Text)
            'str3 = Mid(str1, int1 + Len(str2) + 1, int2 - (int1 + 1))

            strPathMDBDest = str3

        End If

  
        'first update tblversion

        Dim dtNow As Date = Now
        Dim maxID As Int64

        '20190205 LEE
        'Frontage having trouble udating rsVer in SQLSever 2016, though testing here does not show that problem.
        'Try using cmd instead
        Try
            rsVer.MoveLast()
            maxID = rsVer.Fields("ID_TBLVERSION").Value
            maxID = maxID + 1
        Catch ex As Exception

        End Try

        'strSQL works with Access and SQL Server
        strSQL = "INSERT INTO TBLVERSION VALUES (" & maxID & "," & ver(1, 1) & "," & ver(2, 1) & "," & ver(3, 1) & ",'" & dtNow & "'," & ver(4, 1) & ");"

        cmd.CommandText = strSQL
        Try
            cmd.Execute()
            str1 = "SUCCESSFUL:" & ChrW(9) & strSQL & ": Version: " & strVerOld & " to " & strVerNew
        Catch ex As Exception
            var1 = var1
            str1 = "UNSUCCESSFUL:" & ChrW(9) & strSQL & ": Version: " & strVerOld & " to " & strVerNew & ": Exception: " & ex.Message
        End Try

        Call WriteInstall(str1)

        'Try
        '    rsVer.MoveLast()
        '    maxID = rsVer.Fields("ID_TBLVERSION").Value
        '    maxID = maxID + 1
        '    rsVer.AddNew()
        '    rsVer.Fields("ID_TBLVERSION").Value = maxID
        '    rsVer.Fields("INT1VERSION").Value = ver(1, 1)
        '    rsVer.Fields("INT2VERSION").Value = ver(2, 1)
        '    rsVer.Fields("INT3VERSION").Value = ver(3, 1)
        '    Try
        '        rsVer.Fields("INT4VERSION").Value = ver(4, 1)
        '    Catch ex As Exception
        '        var1 = var1
        '    End Try

        '    rsVer.Fields("DTDATE").Value = dtNow
        '    rsVer.Update()
        '    rsVer.ActiveConnection = conA
        '    rsVer.UpdateBatch(AffectEnum.adAffectAllChapters)

        '    If rsVer.State = ADODB.ObjectStateEnum.adStateOpen Then
        '        rsVer.Close()
        '        rsVer = Nothing
        '    End If

        '    '20190131 LEE:
        '    str1 = "SUCCESSFUL:" & ChrW(9) & "rsVer.UpdateBatch: TBLVERSION: Update " & strVerOld & " to " & strVerNew

        'Catch ex As Exception
        '    var1 = var1 'debug
        '    str1 = "UNSUCCESSFUL:" & ChrW(9) & "rsVer.UpdateBatch: TBLVERSION: Update " & strVerOld & " to " & strVerNew & ": MaxID: " & maxID & ": Exception: " & ex.Message
        'End Try

        'Call WriteInstall(str1)

        'skip all this
        'doesn't need to be run anymore

        GoTo skip1

        'open tblTables
        strT = "TBLTABLES"
        rsTables.CursorLocation = CursorLocationEnum.adUseClient
        rsTables.Open("TBLTABLES", conB, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockBatchOptimistic)
        rsTables.ActiveConnection = Nothing
        intRec = rsTables.RecordCount
        Me.pgTable.Maximum = intRec

        Dim boolSkipMaxID As Boolean = False
        Dim boolSkip As Boolean = False
        Dim strlblTable As String
        Dim strlblIndex As String
        Dim strlblInsert As String

        strlblTable = "Evaluating table: "

        'determine if tblMaxID exists in destination
        strT = "TBLMAXID"
        str1 = "SELECT * FROM " & strT
        rsMaxID.CursorLocation = CursorLocationEnum.adUseClient
        Try
            rsMaxID.Open(str1, conA, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockBatchOptimistic)
            rsMaxID.ActiveConnection = Nothing
            boolSkipMaxID = True
        Catch ex As Exception
            boolSkipMaxID = False
        End Try

        If boolSkipMaxID = False Then 'create tblMaxID from conB
            strT = "TBLMAXID"
            str2 = ""
            If boolAccess Then
                str1 = "CREATE TABLE " & strT & ";"
                'str2 = "(ID_TBLVERSION int CONSTRAINT PK_ID_TBLVERSION PRIMARY KEY, INT1VERSION short, INT2VERSION short, INT3VERSION short, DTDATE date);"
                str2 = ""
            ElseIf boolOra Then
                str1 = "CREATE TABLE " & strT & " "
                str2 = "(ID_TBLMAXID NUMBER DEFAULT ( 0 ) CONSTRAINT PK_" & strT & " PRIMARY KEY);" ', INT1VERSION NUMBER DEFAULT ( 0 ), INT2VERSION NUMBER DEFAULT ( 0 ), INT3VERSION NUMBER DEFAULT ( 0 ), DTDATE DATE);"
            ElseIf boolSQLSer Then
            End If
            strSQL = str1 & str2

            myCommand.CommandText = strSQL
            myCommand.ExecuteNonQuery()

            str1 = "SELECT * FROM " & strT & ";"
            rsMaxID.Open(str1, conB, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockBatchOptimistic) 'open source
            rsMaxID.ActiveConnection = Nothing

            'now create fields
            Me.lblTable.Text = strlblTable & strT
            Me.lblTable.Refresh()

            Call Pause(1) 'pause for 1 second

            Call CreateFieldsAccess(boolAccess, boolSQLSer, boolOra, rsMaxID, strT, conA, myCommand, False)
            'now enter records
            Me.pgIndex.Maximum = 1
            Me.pgIndex.Value = 1
            Me.pgIndex.Refresh()
            Call Pause(2)
            Call InsertRecordsAccess(boolAccess, boolSQLSer, boolOra, rsMaxID, strT, conA, myCommand)

            boolSkipMaxID = False
        End If

        boolSkipMaxID = True
        rsTables.MoveFirst()
        cmd.ActiveConnection = conA

        Dim boolC As Boolean

        'If boolAccess Then
        '    dbDAOS = ws.OpenDatabase(strPathDBmdb, False, True, Type.Missing)
        '    dbDAOD = ws.OpenDatabase(strPathMDBDest, False, False, Type.Missing)
        'End If

        'Dim dd As Oracle.DataAccess.Client.OracleConnection

        Call Pause(1)

        Dim arrTableMade(500) As Boolean
        Dim intctT As Short

        intctT = 0

        Me.panCreate.Enabled = True
        Me.panIndex.Enabled = False
        Me.panInsert.Enabled = False
        Me.pgTable.Value = 0
        Me.pgCreate.Value = 0
        Me.pgIndex.Value = 0
        Me.pgInsert.Value = 0

        'don't do this anymore

        GoTo skip1

        Do Until rsTables.EOF

            intctT = intctT + 1
            arrTableMade(intctT) = True
            Me.pgTable.Value = intctT
            Me.pgTable.Refresh()

            Try
                rsD.Close()
            Catch ex As Exception

            End Try
            rsD.CursorLocation = CursorLocationEnum.adUseClient

            strT = rsTables.Fields("CHARTABLENAME").Value
            boolExact = True
            boolSkip = False
            boolMultPK = False
            boolSpecial = False
            boolC = True
            'special: tables have multiple primary keys
            'tblReportStatements
            'TBLTEMPLATEATTRIBUTES
            'tblSummaryData
            'TBLREPORTTABLEANALYTES
            'TBLMETHODVALIDATIONDATA
            'TBLCORPORATEADDRESSES
            'TBLANALYTICALRUNSUMMARY

            Select Case strT
                Case Is = "TBLADDRESSLABELS"
                    boolExact = True
                    'boolSkip = True
                Case Is = "TBLANALREFSTANDARDS"
                    boolExact = False
                    'boolSkip = True
                Case Is = "TBLANALYTICALRUNSUMMARY"
                    boolExact = False
                    boolSpecial = True
                    'boolSkip = True
                Case Is = "TBLAPPFIGS"
                    boolExact = False
                    'boolSkip = True
                Case Is = "TBLASSIGNEDSAMPLES"
                    boolExact = False
                    'boolSkip = True
                Case Is = "TBLASSIGNEDSAMPLESHELPER"
                    boolExact = True
                    'boolSkip = True
                Case Is = "TBLCONFIGAPPFIGS"
                    boolExact = True
                    'boolSkip = True
                Case Is = "TBLCONFIGBODYSECTIONS"
                    boolExact = True
                    'boolSkip = True
                Case Is = "TBLCONFIGHEADERLOOKUP"
                    boolExact = True
                    'boolSkip = True
                Case Is = "TBLCONFIGREPORTTABLES"
                    boolExact = True
                Case Is = "TBLCONFIGREPORTTYPE"
                    boolExact = True
                Case Is = "TBLCONFIGURATION"
                    boolExact = True
                Case Is = "TBLCONTRIBUTINGPERSONNEL"
                    boolExact = False
                Case Is = "TBLCORPORATEADDRESSES"
                    boolExact = False
                    boolSpecial = True
                Case Is = "TBLCORPORATENICKNAMES"
                    boolExact = False
                Case Is = "TBLDATA"
                    boolExact = False
                Case Is = "TBLDATATABLEROWTITLES"
                    boolExact = True
                Case Is = "TBLDATEFORMATS"
                    boolExact = True
                Case Is = "TBLDROPDOWNBOXCONTENT"
                    boolExact = False
                Case Is = "TBLDROPDOWNBOXNAME"
                    boolExact = False
                Case Is = "TBLFIELDCODES"
                    boolExact = True
                Case Is = "TBLHOOKS"
                    boolExact = True
                Case Is = "TBLINCLUDEDROWS"
                    boolExact = False
                Case Is = "TBLMAXID"
                    boolSkip = boolSkipMaxID
                    boolExact = True
                Case Is = "TBLMETHODVALIDATIONDATA"
                    boolExact = False
                    boolSpecial = True
                Case Is = "TBLOUTSTANDINGITEMS"
                    boolExact = False
                Case Is = "TBLPASSWORDHISTORY"
                    boolExact = False
                Case Is = "TBLPERMISSIONS"
                    boolExact = False
                    'boolSkip = True
                Case Is = "TBLPERSONNEL"
                    boolExact = False
                Case Is = "TBLQATABLES"
                    boolExact = False
                Case Is = "TBLREPORTHEADERS"
                    boolExact = False
                Case Is = "TBLREPORTHISTORY"
                    boolExact = False
                Case Is = "TBLREPORTS"
                    boolExact = False
                Case Is = "TBLREPORTSTATEMENTS"
                    boolExact = False
                    boolSpecial = True
                Case Is = "TBLREPORTTABLE"
                    boolExact = False
                Case Is = "TBLREPORTTABLEANALYTES"
                    boolExact = False
                    boolSpecial = True
                Case Is = "TBLREPORTTABLEHEADERCONFIG"
                    boolExact = False
                Case Is = "TBLSAMPLERECEIPT"
                    boolExact = False
                Case Is = "TBLSTUDIES"
                    boolExact = False
                Case Is = "TBLSUMMARYDATA"
                    boolExact = False
                    boolSpecial = True
                Case Is = "TBLTAB1"
                    boolExact = True
                Case Is = "TBLTABLELEGENDS"
                    boolExact = False
                Case Is = "TBLTABLEPROPERTIES"
                    boolExact = False
                Case Is = "TBLTEMPLATEATTRIBUTES"
                    boolExact = False
                    boolSpecial = True
                Case Is = "TBLTEMPLATES"
                    boolExact = False
                Case Is = "TBLUSERACCOUNTS"
                    boolExact = False
                Case Is = "TBLVERSION"
                    boolSkip = True
                    boolExact = False
                Case Is = "TBLWORDDOCS"
                    boolExact = False
                Case Is = "TBLWORDSTATEMENTS"
                    boolExact = False
                Case Is = "TBLWORDSTATEMENTSVERSION"
                    boolExact = False
                Case Is = "TBLSECTIONTEMPLATES"
                    boolExact = False
            End Select

            Me.lblTable.Text = strlblTable & strT
            Me.pgCreate.Value = 0
            Me.pan1.Refresh()

            If boolSkip Then
                arrTableMade(intctT) = True
            Else

                'first determine if table exists in destination
                Try
                    str1 = "SELECT * FROM " & strT & ";"
                    If rsD.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rsD.Close()
                    End If
                    rsD.CursorLocation = CursorLocationEnum.adUseClient
                    rsD.Open(str1, conA, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockBatchOptimistic)
                    rsD.ActiveConnection = Nothing
                    arrTableMade(intctT) = True

                Catch ex As Exception

                    arrTableMade(intctT) = False

                    If boolAccess Then
                        str1 = "CREATE TABLE " & strT & ";"
                    ElseIf boolSQLSer Then
                    ElseIf boolOra Then
                        If boolSpecial Then
                            str1 = "" ' "CREATE TABLE " & strT & " (ID_" & strT & " NUMBER DEFAULT(1));"
                        Else
                            'str1 = "CREATE TABLE " & strT & " (ID_" & strT & " NUMBER DEFAULT(1) PRIMARY KEY)"
                            str1 = "CREATE TABLE " & strT & " (ID_" & strT & " NUMBER DEFAULT(1) CONSTRAINT PK_" & strT & " PRIMARY KEY)"
                            'ADD CONSTRAINT "PK_TBLANALREFSTANDARDS"
                            'str1 = "CREATE TABLE " & strT & " (ID_" & strT & " NUMBER DEFAULT(1));"
                        End If
                    End If
                    If Len(str1) = 0 Then
                    Else
                        myCommand.CommandType = CommandType.Text
                        myCommand.CommandText = str1
                        Try
                            myCommand.ExecuteNonQuery()
                            boolC = True
                        Catch ex1 As Exception
                            boolC = False
                            var1 = ex1.Message.ToString
                            '''''''''''console.writeline(var1)
                            var3 = var1
                        End Try

                        If boolC Then
                            'this is actually source
                            Pause(1)
                            str1 = "SELECT * FROM " & strT & ";"
                            If rsS.State = ADODB.ObjectStateEnum.adStateOpen Then
                                rsS.Close()
                            End If
                            rsS.CursorLocation = CursorLocationEnum.adUseClient
                            rsS.Open(str1, conB, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly) 'open source
                            rsS.ActiveConnection = Nothing

                            Pause(1)
                            Call CreateFieldsAccess(boolAccess, boolSQLSer, boolOra, rsS, strT, conA, myCommand, False)

                            If rsS.State = ADODB.ObjectStateEnum.adStateOpen Then
                                rsS.Close()
                            End If

                        End If

                    End If

                End Try

            End If

            Try
                If rsD.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsD.Close()
                End If
                rsD.ActiveConnection = Nothing
            Catch ex As Exception

            End Try

            Try
                If rsS.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsS.Close()
                End If
                rsS.ActiveConnection = Nothing
            Catch ex As Exception

            End Try

            rsTables.MoveNext()
        Loop

        'GoTo skip1

        Me.pgOverall.Value = 1
        Me.pgOverall.Refresh()

        Me.pgCreate.Value = Me.pgCreate.Maximum
        Me.pgCreate.Refresh()

        Me.panCreate.Enabled = False
        Me.panIndex.Enabled = True
        Me.panInsert.Enabled = False

        'now check and set multiple primary keys
        'If boolAccess Then
        '    dbDAOS = ws.OpenDatabase(strPathDBmdb, False, True, Type.Missing)
        '    dbDAOD = ws.OpenDatabase(strPathMDBDest, False, False, Type.Missing)
        'End If

        boolSkipMaxID = True
        rsTables.MoveFirst()

        Call Pause(1)

        intTNum = 0
        intctT = 0

        Me.pgTable.Value = 0

        Do Until rsTables.EOF
            intctT = intctT + 1

            Me.pgTable.Value = intctT
            Me.pgTable.Refresh()

            If arrTableMade(intctT) Then 'skip
            Else
                Try
                    rsD.Close()
                Catch ex As Exception

                End Try
                rsD.CursorLocation = CursorLocationEnum.adUseClient

                strT = rsTables.Fields("CHARTABLENAME").Value
                boolExact = True
                boolSkip = False
                boolMultPK = False
                boolSpecial = False
                'special: tables have multiple primary keys
                'tblReportStatements
                'TBLTEMPLATEATTRIBUTES
                'tblSummaryData
                'TBLREPORTTABLEANALYTES
                'TBLMETHODVALIDATIONDATA
                'TBLCORPORATEADDRESSES
                'TBLANALYTICALRUNSUMMARY

                intTNum = intTNum + 1
                Select Case strT
                    Case Is = "TBLADDRESSLABELS"
                        boolExact = True
                    Case Is = "TBLANALREFSTANDARDS"
                        boolExact = False
                    Case Is = "TBLANALYTICALRUNSUMMARY"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLAPPFIGS"
                        boolExact = False
                    Case Is = "TBLASSIGNEDSAMPLES"
                        boolExact = False
                    Case Is = "TBLASSIGNEDSAMPLESHELPER"
                        boolExact = True
                    Case Is = "TBLCONFIGAPPFIGS"
                        boolExact = True
                    Case Is = "TBLCONFIGBODYSECTIONS"
                        boolExact = True
                    Case Is = "TBLCONFIGHEADERLOOKUP"
                        boolExact = True
                    Case Is = "TBLCONFIGREPORTTABLES"
                        boolExact = True
                    Case Is = "TBLCONFIGREPORTTYPE"
                        boolExact = True
                    Case Is = "TBLCONFIGURATION"
                        boolExact = True
                    Case Is = "TBLCONTRIBUTINGPERSONNEL"
                        boolExact = False
                    Case Is = "TBLCORPORATEADDRESSES"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLCORPORATENICKNAMES"
                        boolExact = False
                    Case Is = "TBLDATA"
                        boolExact = False
                    Case Is = "TBLDATATABLEROWTITLES"
                        boolExact = True
                    Case Is = "TBLDATEFORMATS"
                        boolExact = True
                    Case Is = "TBLDROPDOWNBOXCONTENT"
                        boolExact = False
                    Case Is = "TBLDROPDOWNBOXNAME"
                        boolExact = False
                    Case Is = "TBLFIELDCODES"
                        boolExact = True
                    Case Is = "TBLHOOKS"
                        boolExact = True
                    Case Is = "TBLINCLUDEDROWS"
                        boolExact = False
                    Case Is = "TBLMAXID"
                        boolSkip = boolSkipMaxID
                        boolExact = True
                    Case Is = "TBLMETHODVALIDATIONDATA"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLOUTSTANDINGITEMS"
                        boolExact = False
                    Case Is = "TBLPASSWORDHISTORY"
                        boolExact = False
                    Case Is = "TBLPERMISSIONS"
                        boolExact = False
                    Case Is = "TBLPERSONNEL"
                        boolExact = False
                    Case Is = "TBLQATABLES"
                        boolExact = False
                    Case Is = "TBLREPORTHEADERS"
                        boolExact = False
                    Case Is = "TBLREPORTHISTORY"
                        boolExact = False
                    Case Is = "TBLREPORTS"
                        boolExact = False
                    Case Is = "TBLREPORTSTATEMENTS"
                        boolExact = False
                        boolSpecial = True

                    Case Is = "TBLREPORTTABLE"
                        boolExact = False
                    Case Is = "TBLREPORTTABLEANALYTES"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLREPORTTABLEHEADERCONFIG"
                        boolExact = False
                    Case Is = "TBLSAMPLERECEIPT"
                        boolExact = False
                    Case Is = "TBLSTUDIES"
                        boolExact = False
                    Case Is = "TBLSUMMARYDATA"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLTAB1"
                        boolExact = True
                    Case Is = "TBLTABLELEGENDS"
                        boolExact = False
                    Case Is = "TBLTABLEPROPERTIES"
                        boolExact = False
                    Case Is = "TBLTEMPLATEATTRIBUTES"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLTEMPLATES"
                        boolExact = False
                    Case Is = "TBLUSERACCOUNTS"
                        boolExact = False
                    Case Is = "TBLVERSION"
                        boolSkip = True
                        boolExact = False
                    Case Is = "TBLWORDDOCS"
                        boolExact = False
                    Case Is = "TBLWORDSTATEMENTS"
                        boolExact = False
                    Case Is = "TBLWORDSTATEMENTSVERSION"
                        boolExact = False
                    Case Is = "TBLSECTIONTEMPLATES"
                        boolExact = False
                End Select

                Me.lblTable.Text = strlblTable & strT
                Me.pgIndex.Value = 0
                Me.pan1.Refresh()

                If boolSkip Then
                Else

                End If

                If rsD.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsD.Close()
                End If
                rsD.ActiveConnection = Nothing
            End If

            rsTables.MoveNext()
        Loop

        'now add data

        Me.pgOverall.Value = 2
        Me.pgOverall.Refresh()

        Me.pgIndex.Value = Me.pgIndex.Maximum
        Me.pgIndex.Refresh()

        Me.panCreate.Enabled = False
        Me.panIndex.Enabled = False
        Me.panInsert.Enabled = True

        Call Pause(1)

        boolSkipMaxID = True
        rsTables.MoveFirst()
        intctT = 0

        Me.pgTable.Value = 0

        Do Until rsTables.EOF

            intctT = intctT + 1
            Me.pgTable.Value = intctT
            Me.pgTable.Refresh()

            If arrTableMade(intctT) Then 'skip
            Else
                Try
                    rsD.Close()
                Catch ex As Exception

                End Try
                rsD.CursorLocation = CursorLocationEnum.adUseClient

                strT = rsTables.Fields("CHARTABLENAME").Value
                boolExact = True
                boolSkip = False
                boolMultPK = False
                boolSpecial = False
                'special: tables have multiple primary keys
                'tblReportStatements
                'TBLTEMPLATEATTRIBUTES
                'tblSummaryData
                'TBLREPORTTABLEANALYTES
                'TBLMETHODVALIDATIONDATA
                'TBLCORPORATEADDRESSES
                'TBLANALYTICALRUNSUMMARY

                Select Case strT
                    Case Is = "TBLADDRESSLABELS"
                        boolExact = True
                    Case Is = "TBLANALREFSTANDARDS"
                        boolExact = False
                    Case Is = "TBLANALYTICALRUNSUMMARY"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLAPPFIGS"
                        boolExact = False
                    Case Is = "TBLASSIGNEDSAMPLES"
                        boolExact = False
                    Case Is = "TBLASSIGNEDSAMPLESHELPER"
                        boolExact = True
                    Case Is = "TBLCONFIGAPPFIGS"
                        boolExact = True
                    Case Is = "TBLCONFIGBODYSECTIONS"
                        boolExact = True
                    Case Is = "TBLCONFIGHEADERLOOKUP"
                        boolExact = True
                    Case Is = "TBLCONFIGREPORTTABLES"
                        boolExact = True
                    Case Is = "TBLCONFIGREPORTTYPE"
                        boolExact = True
                    Case Is = "TBLCONFIGURATION"
                        boolExact = True
                    Case Is = "TBLCONTRIBUTINGPERSONNEL"
                        boolExact = False
                    Case Is = "TBLCORPORATEADDRESSES"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLCORPORATENICKNAMES"
                        boolExact = False
                    Case Is = "TBLDATA"
                        boolExact = False
                    Case Is = "TBLDATATABLEROWTITLES"
                        boolExact = True
                    Case Is = "TBLDATEFORMATS"
                        boolExact = True
                    Case Is = "TBLDROPDOWNBOXCONTENT"
                        boolExact = False
                    Case Is = "TBLDROPDOWNBOXNAME"
                        boolExact = False
                    Case Is = "TBLFIELDCODES"
                        boolExact = True
                    Case Is = "TBLHOOKS"
                        boolExact = True
                    Case Is = "TBLINCLUDEDROWS"
                        boolExact = False
                    Case Is = "TBLMAXID"
                        boolSkip = boolSkipMaxID
                        boolExact = True
                    Case Is = "TBLMETHODVALIDATIONDATA"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLOUTSTANDINGITEMS"
                        boolExact = False
                    Case Is = "TBLPASSWORDHISTORY"
                        boolExact = False
                    Case Is = "TBLPERMISSIONS"
                        boolExact = False
                    Case Is = "TBLPERSONNEL"
                        boolExact = False
                    Case Is = "TBLQATABLES"
                        boolExact = False
                    Case Is = "TBLREPORTHEADERS"
                        boolExact = False
                    Case Is = "TBLREPORTHISTORY"
                        boolExact = False
                    Case Is = "TBLREPORTS"
                        boolExact = False
                    Case Is = "TBLREPORTSTATEMENTS"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLREPORTTABLE"
                        boolExact = False
                    Case Is = "TBLREPORTTABLEANALYTES"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLREPORTTABLEHEADERCONFIG"
                        boolExact = False
                    Case Is = "TBLSAMPLERECEIPT"
                        boolExact = False
                    Case Is = "TBLSTUDIES"
                        boolExact = False
                    Case Is = "TBLSUMMARYDATA"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLTAB1"
                        boolExact = True
                    Case Is = "TBLTABLELEGENDS"
                        boolExact = False
                    Case Is = "TBLTABLEPROPERTIES"
                        boolExact = False
                    Case Is = "TBLTEMPLATEATTRIBUTES"
                        boolExact = False
                        boolSpecial = True
                    Case Is = "TBLTEMPLATES"
                        boolExact = False
                    Case Is = "TBLUSERACCOUNTS"
                        boolExact = False
                    Case Is = "TBLVERSION"
                        boolSkip = True
                        boolExact = False
                    Case Is = "TBLWORDDOCS"
                        boolExact = False
                    Case Is = "TBLWORDSTATEMENTS"
                        boolExact = False

                End Select

                Me.lblTable.Text = strlblTable & strT
                Me.pgInsert.Value = 0
                Me.pan1.Refresh()

                If boolSkip Then
                Else

                    'first determine if table exists

                    Try
                        str1 = "SELECT * FROM " & strT & ";"
                        If rsD.State = ADODB.ObjectStateEnum.adStateOpen Then
                            rsD.Close()
                        End If
                        rsD.Open(str1, conB, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockBatchOptimistic)

                        Pause(1)
                        Call InsertRecordsAccess(boolAccess, boolSQLSer, boolOra, rsD, strT, conA, myCommand)
                        'conA.CommitTrans()

                    Catch ex As Exception

                        var1 = "aa" 'debug
                    End Try

                End If

                If rsD.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsD.Close()
                End If
                rsD.ActiveConnection = Nothing
            End If



            rsTables.MoveNext()
        Loop

skip1:

        Me.pgOverall.Maximum = 75

        'now dow individual table mods
        Dim intCount1 As Short
        intCount1 = 0
        Try
            strNV_App = Format(ver(1, 1), "00") & Format(ver(2, 1), "00") & Format(ver(3, 1), "00")
            strNV_DB = Format(ver(1, 2), "00") & Format(ver(2, 2), "00") & Format(ver(3, 2), "00")
            strNV_DB_New = Format(ver(1, 2), "00") & Format(ver(2, 2), "00") & Format(ver(3, 2), "00") & Format(ver(4, 2), "00")

            Call DoIndUpdates(Me, boolAccess, boolSQLSer, boolOra, conB, myCommand, constrB, Me.pgOverall)

            str1 = "SUCCESSFUL:" & ChrW(9) & "DoIndUpdates" '20190131 LEE:

        Catch ex As Exception
            MsgBox("Error..." & ChrW(10) & ex.Message)
            str1 = "UNSUCCESSFUL" & ChrW(9) & "DoIndUpdates: Exception: " & ex.Message '20190131 LEE:
        End Try

        '20190131 LEE:
        Call WriteInstall(str1)
        Call WriteInstallDir()

        Me.pgOverall.Value = Me.pgOverall.Maximum
        Me.pgOverall.Refresh()

        Me.pgInsert.Value = Me.pgInsert.Maximum
        Me.pgInsert.Refresh()

        'now that all tables exist, may do specific version related updates


        If rsS.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsS.Close()
        End If
        rsS = Nothing

        If rsD.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsD.Close()
        End If
        rsD = Nothing

        If rsMaxID.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsMaxID.Close()
        End If
        rsMaxID = Nothing

        If rsTables.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsTables.Close()
        End If
        rsTables = Nothing


    End Sub

    Sub CreateFieldsAccess(ByVal boolAccess As Boolean, ByVal boolSQLSer As Boolean, ByVal boolOra As Boolean, ByVal rsS As ADODB.Recordset, ByVal strT As String, ByVal conA As ADODB.Connection, ByVal cmd As OleDb.OleDbCommand, ByVal boolTExists As Boolean)

        Dim fldS As ADODB.Field
        Dim var1, var2, var3, var4, var5, var6
        Dim fldprop As ADODB.Properties
        Dim Count1 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strID As String
        Dim strTy As String
        Dim varType
        Dim varSize
        Dim boolSpecial As Boolean
        Dim boolHasISStudies As Boolean = False
        Dim strSQL As String
        Dim strFN As String

        Dim pgV As Int16
        Dim pgMax As Int16
        Dim intCt As Short

        intCt = rsS.Fields.Count
        pgMax = intCt
        Me.pgCreate.Value = 0
        Me.pgCreate.Maximum = intCt
        Me.pgCreate.Refresh()

        'adInteger;3;ID_TBLADDRESSLABELS
        'adVarWChar;202;CHARLABEL
        'adDate;7;UPSIZE_TS
        'adSmallInt;2;BOOLISREPLICATE
        'adDouble;5;NUMWATSONRUNNUMBER
        'adLongVarWChar;203;CHARFIELDCODE

        'conA.BeginTrans()

        'first create fields all as number
        Count1 = 0
        For Each fldS In rsS.Fields
            Count1 = Count1 + 1
            Me.pgCreate.Value = Count1
            Me.pgCreate.Refresh()

            var1 = fldS.Name
            strFN = var1
            var2 = fldS.Type
            var3 = fldS.Attributes
            var4 = fldS.DefinedSize
            var5 = fldS.Precision
            'var6 = fldS.ActualSize

            strTy = ""

            If Count1 = 1 Then
                strID = strFN
            Else
                'look for id_tblstudies
                If StrComp(strFN, "ID_TBLSTUDIES", CompareMethod.Text) = 0 Then
                    boolHasISStudies = True
                End If
            End If



            If Count1 = 1 And boolTExists = False Then
                strTy = "AA"
                If StrComp(strT, "tblReportStatements", vbTextCompare) = 0 Then
                    'Call Special_tblReportStatements(tdfNew)
                    boolSpecial = True
                ElseIf StrComp(strT, "TBLTEMPLATEATTRIBUTES", vbTextCompare) = 0 Then
                    'Call Special_TBLTEMPLATEATTRIBUTES(tdfNew)
                    boolSpecial = True
                ElseIf StrComp(strT, "tblSummaryData", vbTextCompare) = 0 Then
                    'Call Special_tblSummaryData(tdfNew)
                    boolSpecial = True
                ElseIf StrComp(strT, "TBLREPORTTABLEANALYTES", vbTextCompare) = 0 Then
                    'Call Special_TBLREPORTTABLEANALYTES(tdfNew)
                    boolSpecial = True
                ElseIf StrComp(strT, "TBLMETHODVALIDATIONDATA", vbTextCompare) = 0 Then
                    'Call Special_TBLMETHODVALIDATIONDATA(tdfNew)
                    boolSpecial = True
                ElseIf StrComp(strT, "TBLCORPORATEADDRESSES", vbTextCompare) = 0 Then
                    'Call Special_TBLCORPORATEADDRESSES(tdfNew)
                    boolSpecial = True
                ElseIf StrComp(strT, "TBLANALYTICALRUNSUMMARY", vbTextCompare) = 0 Then
                    'Call Special_TBLANALYTICALRUNSUMMARY(tdfNew)
                    boolSpecial = True
                Else
                    boolSpecial = False
                End If

                If boolAccess Then
                    If boolSpecial Then
                        str1 = "alter table " & strT & " ADD " & strFN & " int" ';"
                    Else
                        str1 = "alter table " & strT & " ADD " & strFN & " int CONSTRAINT PK_" & strT & " PRIMARY KEY" ';"
                        'str1 = "alter table " & strT & " ADD " & strfn & " int PRIMARY KEY"';"
                    End If
                    strTy = "BB"
                ElseIf boolSQLSer Then
                ElseIf boolOra Then
                    If boolSpecial Then 'need to create a table
                        str1 = "CREATE TABLE " & strT & " "
                        Select Case var2 'NO SEMICOLONS HERE!!
                            Case Is = 3 'adInteger;3;ID_TBLADDRESSLABELS
                                strTy = "NUMBER DEFAULT(0)"
                            Case Is = 202 'adVarWChar;202;CHARLABEL
                                strTy = "VARCHAR2(4000)"
                            Case Is = 7 'adDate;7;UPSIZE_TS
                                strTy = "DATE"
                            Case Is = 2 'adSmallInt;2;BOOLISREPLICATE
                                strTy = "NUMBER DEFAULT(0)"
                            Case Is = 5 'adDouble;5;NUMWATSONRUNNUMBER
                                strTy = "NUMBER DEFAULT(0)"
                            Case Is = 203 'adLongVarWChar;203;CHARFIELDCODE
                                strTy = "VARCHAR2(4000)"
                        End Select
                        str2 = "(" & strFN & " " & strTy & ")" ';"
                        strSQL = str1 & str2
                        str1 = strSQL
                    Else 'already created
                    End If

                End If
            Else
                If boolAccess Then
                    Select Case var2
                        Case Is = 3 'adInteger;3;ID_TBLADDRESSLABELS
                            strTy = "int NULL" ';"
                        Case Is = 202 'adVarWChar;202;CHARLABEL
                            strTy = "text(" & var4 & ")" ';"
                        Case Is = 7 'adDate;7;UPSIZE_TS
                            strTy = "date NULL" ';"
                        Case Is = 2 'adSmallInt;2;BOOLISREPLICATE
                            strTy = "short NULL" ';"
                        Case Is = 5 'adDouble;5;NUMWATSONRUNNUMBER
                            strTy = "double NULL" ';"
                        Case Is = 203 'adLongVarWChar;203;CHARFIELDCODE
                            strTy = "memo NULL" ';"
                    End Select
                ElseIf boolSQLSer Then
                ElseIf boolOra Then
                    Select Case var2
                        Case Is = 3 'adInteger;3;ID_TBLADDRESSLABELS
                            strTy = "NUMBER DEFAULT(0)" ';"
                        Case Is = 202 'adVarWChar;202;CHARLABEL
                            strTy = "VARCHAR2(4000)" ';"
                        Case Is = 7 'adDate;7;UPSIZE_TS
                            strTy = "DATE" ';"
                        Case Is = 2 'adSmallInt;2;BOOLISREPLICATE
                            strTy = "NUMBER DEFAULT(0)" ';"
                        Case Is = 5 'adDouble;5;NUMWATSONRUNNUMBER
                            strTy = "NUMBER DEFAULT(0)" ';"
                        Case Is = 203 'adLongVarWChar;203;CHARFIELDCODE
                            strTy = "VARCHAR2(4000)" ';"
                    End Select
                End If

                str1 = "alter table " & strT & " ADD " & strFN & " " & strTy
            End If
            If Len(strTy) = 0 Or StrComp(strTy, "AA", CompareMethod.Text) = 0 Then
            Else
                cmd.CommandText = str1
                Try
                    cmd.ExecuteNonQuery()
                    var2 = "Hi"
                Catch ex As Exception
                    var1 = ex.Message.ToString
                    var2 = var1
                End Try
            End If

        Next

        'conA.CommitTrans()

    End Sub

    'Sub MakeMultPKeys(ByVal tdfnew As dao.TableDef, ByVal boolAccess As Boolean, ByVal boolSQLSer As Boolean, ByVal boolOra As Boolean, ByVal rsS As ADODB.Recordset, ByVal strT As String, ByVal conA As ADODB.Connection, ByVal cmd As OleDb.OleDbCommand, ByVal intTNum As Short)

    '    ''Dim tdfnew As dao.TableDef
    '    'Dim tdfnew As New dao.TableDef
    '    Dim boolSpecial As Boolean = False
    '    Dim strPK As String
    '    Dim str1 As String
    '    Dim str2 As String
    '    Dim str3 As String
    '    Dim str4 As String

    '    Dim Count1 As Short
    '    Dim ind As dao.Index
    '    Dim strTNum As String
    '    Dim fldS As ADODB.Field
    '    Dim var1, var2

    '    strTNum = Format(intTNum, "000")

    '    If boolAccess Then
    '        If StrComp(strT, "tblReportStatements", vbTextCompare) = 0 Then
    '            Call Special_tblReportStatements(tdfnew)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLTEMPLATEATTRIBUTES", vbTextCompare) = 0 Then
    '            Call Special_TBLTEMPLATEATTRIBUTES(tdfnew)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "tblSummaryData", vbTextCompare) = 0 Then
    '            Call Special_tblSummaryData(tdfnew)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLREPORTTABLEANALYTES", vbTextCompare) = 0 Then
    '            Call Special_TBLREPORTTABLEANALYTES(tdfnew)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLMETHODVALIDATIONDATA", vbTextCompare) = 0 Then
    '            Call Special_TBLMETHODVALIDATIONDATA(tdfnew)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLCORPORATEADDRESSES", vbTextCompare) = 0 Then
    '            Call Special_TBLCORPORATEADDRESSES(tdfnew)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLANALYTICALRUNSUMMARY", vbTextCompare) = 0 Then
    '            Call Special_TBLANALYTICALRUNSUMMARY(tdfnew)
    '            boolSpecial = True
    '        Else
    '            'create index for id_tblstudies
    '            If StrComp(strT, "TBLSTUDIES", vbTextCompare) = 0 Then
    '            Else
    '                strPK = "ID_TBLSTUDIES"
    '                For Count1 = 0 To tdfnew.Fields.Count - 1
    '                    str1 = tdfnew.Fields(Count1).Name
    '                    If StrComp(strPK, str1, vbTextCompare) = 0 Then 'make index
    '                        ind = tdfnew.CreateIndex(strPK)
    '                        ind.Fields.Append(ind.CreateField(strPK))
    '                        tdfnew.Indexes.Append(ind)
    '                        Exit For
    '                    End If
    '                Next
    '            End If
    '        End If
    '    ElseIf boolSQLSer Then
    '    ElseIf boolOra Then
    '        If StrComp(strT, "tblReportStatements", vbTextCompare) = 0 Then
    '            Call Special_tblReportStatementsORA(cmd, strT)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLTEMPLATEATTRIBUTES", vbTextCompare) = 0 Then
    '            Call Special_TBLTEMPLATEATTRIBUTESORA(cmd, strT)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "tblSummaryData", vbTextCompare) = 0 Then
    '            Call Special_tblSummaryDataORA(cmd, strT)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLREPORTTABLEANALYTES", vbTextCompare) = 0 Then
    '            Call Special_TBLREPORTTABLEANALYTESORA(cmd, strT)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLMETHODVALIDATIONDATA", vbTextCompare) = 0 Then
    '            Call Special_TBLMETHODVALIDATIONDATAORA(cmd, strT)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLCORPORATEADDRESSES", vbTextCompare) = 0 Then
    '            Call Special_TBLCORPORATEADDRESSESORA(cmd, strT)
    '            boolSpecial = True
    '        ElseIf StrComp(strT, "TBLANALYTICALRUNSUMMARY", vbTextCompare) = 0 Then
    '            Call Special_TBLANALYTICALRUNSUMMARYORA(cmd, strT)
    '            boolSpecial = True
    '        Else
    '            'create index for id_tblstudies
    '            If StrComp(strT, "TBLSTUDIES", vbTextCompare) = 0 Then
    '            Else
    '                strPK = "ID_TBLSTUDIES"
    '                For Each fldS In rsS.Fields
    '                    str1 = fldS.Name
    '                    If StrComp(strPK, str1, vbTextCompare) = 0 Then 'make index
    '                        str1 = "CREATE INDEX ID_TBLSTUDIES_" & strTNum & " ON " & strT & " (ID_TBLSTUDIES)"
    '                        cmd.CommandText = str1
    '                        Try
    '                            cmd.ExecuteNonQuery()
    '                        Catch ex As Exception
    '                            var1 = ex.Message
    '                            var2 = var1
    '                        End Try
    '                        Exit For
    '                    End If

    '                Next
    '            End If
    '        End If
    '    End If

    'End Sub

    Sub InsertRecordsAccess(ByVal boolAccess As Boolean, ByVal boolSQLSer As Boolean, ByVal boolOra As Boolean, ByVal rsS As ADODB.Recordset, ByVal strT As String, ByVal conA As ADODB.Connection, ByVal cmd As OleDb.OleDbCommand)
        Dim fldS As ADODB.Field
        Dim fldS1 As ADODB.Field
        Dim rsD As New ADODB.Recordset
        Dim fldD As ADODB.Field
        Dim var1, var2, var3, var4, var5, var6, var7, var5a
        Dim str1 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim id As Int16
        Dim strID As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim varE1, varE2
        Dim intS As Int64
        Dim intE As Int64

        Dim pgV As Int16
        Dim pgMax As Int16
        Dim intCt As Short
        Dim intCt1 As Int16

        intCt = rsS.RecordCount
        pgMax = intCt
        Me.pgInsert.Value = 0
        Me.pgInsert.Maximum = intCt
        Me.pgInsert.Refresh()

        Dim boolWdDoc As Boolean = False

        Try
            If StrComp(strT, "TBLWORDDOCS", CompareMethod.Text) = 0 Then
                boolWdDoc = True
            End If

            If rsS.EOF And rsS.BOF Then 'no records to add

            Else
                'now open recordset
                rsD.CursorLocation = CursorLocationEnum.adUseClient
                str1 = "SELECT * FROM " & strT
                rsD.Open(str1, conA, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockBatchOptimistic)
                rsD.ActiveConnection = Nothing

                int4 = 0
                If rsS.EOF And rsS.BOF Then
                Else
                    rsS.MoveFirst()
                    int1 = 0
                    int2 = 3
                    intCt1 = 0
                    Do Until rsS.EOF
                        intCt1 = intCt1 + 1
                        Me.pgInsert.Value = intCt1
                        Me.pgInsert.Refresh()

                        rsD.AddNew()
                        int3 = 0
                        For Each fldS In rsS.Fields
                            int3 = int3 + 1
                            'If int3 = 1 Then
                            '    id = fldS.Value + int4
                            '    strID = fldS.Name
                            'End If
                            Try
                                'var3 = fldS.ActualSize
                                'var4 = fldS.DefinedSize
                                var5 = Len(fldS.Value)
                                var6 = rsD.Fields(fldS.Name).Status
                                'If int3 = 1 Then
                                '    rsD.Fields(fldS.Name).Value = NZ(fldS.Value + int4, System.DBNull.Value)
                                'Else
                                '    rsD.Fields(fldS.Name).Value = NZ(fldS.Value, System.DBNull.Value)
                                'End If
                                rsD.Fields(fldS.Name).Value = NZ(fldS.Value, System.DBNull.Value)
                            Catch ex As Exception
                                var1 = ex.Message
                                If InStr(1, var1, "Type", CompareMethod.Text) > 0 Then
                                Else
                                    var7 = rsD.Fields(fldS.Name).Status
                                    'If var7 = 2 Then 'truncate data
                                    '    int4 = int4 + 1
                                    '    intS = CInt(Len(fldS.Value) / 2)
                                    '    intE = Len(fldS.Value) - intS
                                    '    varE1 = Mid(fldS.Value, 1, intS)
                                    '    varE2 = Mid(fldS.Value, intS + 1, intE)
                                    '    Try
                                    '        var3 = Len(varE1)
                                    '        rsD.Fields(fldS.Name).Value = varE1
                                    '    Catch ex1 As Exception
                                    '        var7 = rsD.Fields(fldS.Name).Status
                                    '        var1 = ex1.Message
                                    '        var2 = var1
                                    '    End Try
                                    '    Try
                                    '        var4 = Len(varE2)
                                    '        rsD.Fields(fldS.Name).Value = varE2
                                    '    Catch ex1 As Exception
                                    '        var7 = rsD.Fields(fldS.Name).Status
                                    '        var1 = ex1.Message
                                    '        var2 = var1
                                    '    End Try

                                    '    'redo data entry
                                    '    For Each fldS1 In rsS.Fields

                                    '    Next
                                    'End If

                                End If
                            End Try
                        Next
                        rsD.Update()
                        If boolWdDoc Then
                            'TBLWORDDOCS has a large amount of data
                            'seems to overwhelm updatebatch unless done in batches
                            int1 = int1 + 1
                            If int1 = int2 Then
                                int2 = int2 + 3
                                rsD.ActiveConnection = conA
                                Try
                                    rsD.UpdateBatch(AffectEnum.adAffectAllChapters)
                                    str1 = "SUCCESSFUL:" & ChrW(9) & "rsD.UpdateBatch: TBLWORDDOCS: Sub InsertRecordsAccess"
                                Catch ex As Exception
                                    str1 = "UNSUCCESSFUL:" & ChrW(9) & "rsD.UpdateBatch: TBLWORDDOCS: Sub InsertRecordsAccess: Error: " & ex.Message
                                    var1 = ex.Message
                                    var2 = var1
                                End Try

                                Call WriteInstall(str1)
                                rsD.ActiveConnection = Nothing
                            End If
                        End If

                        rsS.MoveNext()
                    Loop
                End If

                rsD.ActiveConnection = conA
                Try
                    rsD.UpdateBatch(AffectEnum.adAffectAllChapters)
                Catch ex As Exception
                    MsgBox("Problem inserting records into " & strT)
                End Try
                rsD.ActiveConnection = Nothing
                rsD.Close()

            End If
        Catch ex1 As Exception
            var1 = ex1.Message
            var2 = var1

        End Try



    End Sub

    Function CreateTableStringOra(ByVal rsS As ADODB.Recordset, ByVal strT As String)
        Dim fld As ADODB.Field
        Dim var1, var2, var3, var4

        For Each fld In rsS.Fields
            var1 = fld.Name
            var2 = fld.Type
            var3 = fld.Attributes

        Next



    End Function

    Function CreateTableStringAccess(ByVal rsS As ADODB.Recordset, ByVal strT As String)

    End Function

    Function SetPropertyDAO(ByVal obj As Object, ByVal strPropertyName As String, ByVal intType As Integer, _
ByVal varValue As Object, Optional ByVal strErrMsg As String = "") As Boolean

        On Error GoTo ErrHandler
        'Purpose:   Set a property for an object, creating if necessary.
        'Arguments: obj = the object whose property should be set.
        '           strPropertyName = the name of the property to set.
        '           intType = the type of property (needed for creating)
        '           varValue = the value to set this property to.
        '           strErrMsg = string to append any error message to.

        If HasProperty(obj, strPropertyName) Then
            obj.Properties(strPropertyName) = varValue
        Else
            obj.Properties.Append(obj.CreateProperty(strPropertyName, intType, varValue))
        End If
        SetPropertyDAO = True

ExitHandler:

        Exit Function

ErrHandler:
        strErrMsg = strErrMsg & obj.Name & "." & strPropertyName & " not set to " & varValue & _
            ". Error " & Err.Number & " - " & Err.Description & vbCrLf

        Resume ExitHandler
    End Function

    Public Function HasProperty(ByVal obj As Object, ByVal strPropName As String) As Boolean
        'Purpose:   Return true if the object has the property.
        Dim varDummy As Object

        On Error Resume Next
        varDummy = obj.Properties(strPropName)
        HasProperty = (Err.Number = 0)
    End Function



    'START
    Sub Special_tblReportStatementsORA(ByVal cmd As OleDb.OleDbCommand, ByVal strT As String)
        Dim str1 As String
        Dim str2 As String

        str1 = "ALTER TABLE " & strT & " ADD CONSTRAINT PK_" & strT & " PRIMARY KEY ("
        str1 = str1 & "ID_TBLSTUDIES, "
        str1 = str1 & "ID_TBLCONFIGREPORTTYPE, "
        str1 = str1 & "ID_TBLCONFIGBODYSECTIONS) "
        str1 = str1 & "ENABLE"
        cmd.CommandText = str1

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try

    End Sub

    Sub Special_TBLTEMPLATEATTRIBUTESORA(ByVal cmd As OleDb.OleDbCommand, ByVal strT As String)
        Dim str1 As String
        Dim str2 As String

        str1 = "ALTER TABLE " & strT & " ADD CONSTRAINT PK_" & strT & " PRIMARY KEY ("
        str1 = str1 & "ID_TBLTEMPLATES, "
        str1 = str1 & "ID_TBLTAB1) "
        str1 = str1 & "ENABLE"
        cmd.CommandText = str1

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try

    End Sub

    Sub Special_tblSummaryDataORA(ByVal cmd As OleDb.OleDbCommand, ByVal strT As String)

        Dim str1 As String
        Dim str2 As String

        str1 = "ALTER TABLE " & strT & " ADD CONSTRAINT PK_" & strT & " PRIMARY KEY ("
        str1 = str1 & "ID_TBLSTUDIES, "
        str1 = str1 & "ID_TBLDATATABLEROWTITLES) "
        str1 = str1 & "ENABLE"
        cmd.CommandText = str1

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try


        'Dim ind As dao.Index

        'ind = tdfNew.CreateIndex("PK_tblSummaryData")
        'With ind
        '    .Fields.Append(ind.CreateField("ID_TBLSTUDIES"))
        '    .Fields.Append(ind.CreateField("ID_TBLDATATABLEROWTITLES"))
        '    '.Unique = False
        '    .Primary = True
        'End With
        'tdfNew.Indexes.Append(ind)

    End Sub
    '
    Sub Special_TBLREPORTTABLEANALYTESORA(ByVal cmd As OleDb.OleDbCommand, ByVal strT As String)

        Dim str1 As String
        Dim str2 As String

        str1 = "ALTER TABLE " & strT & " ADD CONSTRAINT PK_" & strT & " PRIMARY KEY ("
        str1 = str1 & "ID_TBLSTUDIES, "
        str1 = str1 & "ANALYTEID, "
        str1 = str1 & "ID_TBLREPORTTABLE, "
        str1 = str1 & "ANALYTEINDEX, "
        str1 = str1 & "MASTERASSAYID) "
        str1 = str1 & "ENABLE"
        cmd.CommandText = str1

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try


    End Sub
    'TBLMETHODVALIDATIONDATA
    Sub Special_TBLMETHODVALIDATIONDATAORA(ByVal cmd As OleDb.OleDbCommand, ByVal strT As String)

        Dim str1 As String
        Dim str2 As String

        str1 = "ALTER TABLE " & strT & " ADD CONSTRAINT PK_" & strT & " PRIMARY KEY ("
        str1 = str1 & "ID_TBLSTUDIES, "
        str1 = str1 & "INTCOLUMNNUMBER) "
        str1 = str1 & "ENABLE"
        cmd.CommandText = str1

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try

    End Sub
    '
    Sub Special_TBLCORPORATEADDRESSESORA(ByVal cmd As OleDb.OleDbCommand, ByVal strT As String)

        Dim str1 As String
        Dim str2 As String

        str1 = "ALTER TABLE " & strT & " ADD CONSTRAINT PK_" & strT & " PRIMARY KEY ("
        str1 = str1 & "ID_TBLCORPORATENICKNAMES, "
        str1 = str1 & "ID_TBLADDRESSLABELS) "
        str1 = str1 & "ENABLE"
        cmd.CommandText = str1

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try

    End Sub
    'TBLANALYTICALRUNSUMMARY
    Sub Special_TBLANALYTICALRUNSUMMARYORA(ByVal cmd As OleDb.OleDbCommand, ByVal strT As String)

        'ALTER TABLE "TBLANALYTICALRUNSUMMARY" ADD CONSTRAINT "PK_TBLANALYTICALRUNSUMMARY" PRIMARY KEY ("ID_TBLSTUDIES", "INTWATSONRUNID", "CHARANALYTE") ENABLE;

        Dim str1 As String
        Dim str2 As String
        Dim var1, var2

        str1 = "ALTER TABLE " & strT & " ADD CONSTRAINT PK_" & strT & " PRIMARY KEY ("
        str1 = str1 & "ID_TBLSTUDIES, "
        str1 = str1 & "INTWATSONRUNID, "
        str1 = str1 & "CHARANALYTE) "
        str1 = str1 & "ENABLE"
        cmd.CommandText = str1

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            var1 = ex.Message
            var2 = var1
        End Try

    End Sub


End Class