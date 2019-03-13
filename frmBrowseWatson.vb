Option Compare Text

Imports System.Data
Imports System.Linq.Expressions
Imports System.Linq
Imports System.Data.DataTableExtensions
Imports System.Data.DataRowExtensions

Public Class frmBrowseWatson

    Public boolCancel As Boolean
    Public boolHold As Boolean = False
    Public boolGetOracle As Boolean
    Public boolArchive As Boolean
    Public boolCompare As Boolean = False
    Public boolFormLoad As Boolean = False

    Private Sub frmBrowseWatson_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        boolCancel = True
        boolHold = False

        boolFormLoad = True

        Call modExtensionMethods.DoubleBufferedControl(Me, True)
        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

        If frmH.rbOracle.Checked Then
            boolArchive = False
        Else
            boolArchive = True
        End If

        Try
            Call LoadGrids()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Call ConfigGrids(True)

        boolFormLoad = False

        If Me.dgvProjects.RowCount = 0 Then
        Else
            Me.dgvProjects.Rows(0).Selected = True
            Call ClickProjects()
        End If

    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.EventArgs) Handles cmdOK.Click

        Dim strM As String
        Dim id1 As Int64
        Dim id2 As Int64
        Dim strf As String
        Dim dtNow As Date = Now

        Dim tUserID As String
        Dim tUserName As String

        Dim dgvP As DataGridView = Me.dgvProjects
        Dim dgvS As DataGridView = Me.dgvStudies

        Dim intRowP As Int32
        Dim intRowS As Int32

        Try
            intRowP = dgvP.CurrentRow.Index
        Catch ex As Exception
            strM = "A Project must be selected."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End Try

        Try
            intRowS = dgvS.CurrentRow.Index
        Catch ex As Exception
            strM = "A Study must be selected."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End Try

        gWID = 0
        gWPID = 0
        boolNewOracle = True

        id1 = Me.txtProjectID.Text
        id2 = Me.txtStudyID.Text

        Dim rows() As DataRow

        'strf = "PROJECTID = " & id1 & " AND STUDYID = " & id2

        'strf = "INT_WATSONID = " & id2
        '' rows = tblStudies.Select(strf)

        strf = "INT_WATSONPROJECTID = " & id1 & " AND INT_WATSONSTUDYID = " & id2
        wStudyID = id2

        ''these two id's need to be saved globally to use when opening new Oracle data
        gWPID = id1
        gWID = id2

        rows = tblStudies.Select(strf)

        If boolGetOracle Then

            If rows.Length = 0 Then 'must add record to tblstudies

                'hide the form at this point
                boolCancel = False
                Me.Visible = False

                'make frmh focused
                frmH.Activate()


                boolNewOracle = True

                'id_tblStudies 44
                'INT_WATSONSTUDYID -1
                'INT_WATSONPROJECTID -1
                'CHARWATSONSTUDYNAME MethodDev
                'DTCONFIGURED 16 - Sep - 7
                'UPSIZE_TS()  <-Skip this
                'CHARCUST()  <-Skip this
                'ID_TBLDATASYSTEM 1
                Dim intMax As Int64 = GetMaxID("TBLSTUDIES", 1, True) 'Note: this also increases maxid

                Dim dr1 As DataRow = tblStudies.NewRow
                dr1.BeginEdit()
                dr1("id_tblStudies") = intMax
                dr1("INT_WATSONSTUDYID") = id2
                dr1("INT_WATSONPROJECTID") = id1
                dr1("CHARWATSONSTUDYNAME") = dgvS("STUDYNAME", intRowS).Value
                dr1("DTCONFIGURED") = dtNow
                dr1("ID_TBLDATASYSTEM") = 1 '=Thermo Watson
                dr1.EndEdit()

                id_tblStudies = intMax

                tblStudies.Rows.Add(dr1)

                'check for audit trail
                If gboolAuditTrail Then

                    'add audit trail
                    'clear audittrailtemp
                    tblAuditTrailTemp.Clear()
                    idSE = 0
                    Call FillAuditTrailTemp(tblStudies)

                End If

                If boolGuWuOracle Then
                    Try
                        'ta_tblMaxID.Update(tblMaxID)
                        ta_tblStudies.Update(tblStudies)
                    Catch ex As DBConcurrencyException
                        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                    End Try
                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblStudiesAcc.Update(tblStudies)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                    End Try
                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblStudiesSQLServer.Update(tblStudies)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                    End Try
                End If

                'record tblaudittrailtemp
                Call RecordAuditTrail(False, dtNow)

                'need to add a record in tblReports
                Dim intMax1 = GetMaxID("TBLREPORTS", 1, True)
                Dim dr2 As DataRow = tblReports.NewRow
                dr2.BeginEdit()

                dr2("ID_TBLREPORTS") = intMax1
                dr2("ID_TBLSTUDIES") = intMax

                'enter some defaults
                dr2("INTCALSTD") = 1
                dr2("INTQC") = 1
                dr2("INTSHOWBQL") = 1
                dr2("INTSHOWCALSTD") = 1
                dr2("INTUSERCOMMENTS") = 2
                dr2("BOOLEXCLUDEPSAE") = 1 'deprecated
                dr2("BOOLMULTIVALSUM") = -1
                dr2("BOOLDISPLAYATTACHMENTS") = 0
                dr2("BOOLINSERTWORDDOCS") = 0
                dr2("BOOLREADONLYTABLES") = 0

                'BOOLALLAR chkAll
                'BOOLACCAR chkAccepted
                'BOOLREJAR chkRejected
                'BOOLREGRAR chkRegrPerformed
                'BOOLNOREGRAR chkNoRegrPerformed
                'BOOLINCLPSAE chkPSAE

                dr2("BOOLALLAR") = -1
                dr2("BOOLACCAR") = 0
                dr2("BOOLREJAR") = 0
                dr2("BOOLREGRAR") = 0
                dr2("BOOLNOREGRAR") = 0
                dr2("BOOLINCLPSAE") = 0

                dr2.EndEdit()
                tblReports.Rows.Add(dr2)

                'don't do audit trail
                If boolGuWuOracle Then
                    Try
                        'ta_tblMaxID.Update(tblMaxID)
                        ta_tblReports.Update(tblReports)
                    Catch ex As DBConcurrencyException
                        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
                    End Try
                ElseIf boolGuWuAccess Then
                    Try
                        ta_tblReportsAcc.Update(tblReports)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                    End Try
                ElseIf boolGuWuSQLServer Then
                    Try
                        ta_tblReportsSQLServer.Update(tblReports)
                    Catch ex As DBConcurrencyException
                        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
                    End Try
                End If


                'at this point, something must trigger getstudyinfo
                Call frmH.OpenArchiveMDB(False)

                'now select appropriate row in dgvw
                Dim dgv As DataGridView = frmH.dgvwStudy
                Dim int1 As Int16 = dgv.Rows.Count
                Dim int2 As Int64
                Dim Count1 As Int16
                For Count1 = 0 To int1 - 1
                    int2 = dgv("STUDYID", Count1).Value
                    If int2 = id2 Then
                        Dim boolTTT As Boolean
                        boolTTT = boolFormLoad
                        boolFormLoad = True
                        dgv.CurrentCell = dgv.Item("STUDYNAME", Count1)
                        dgv.Rows(Count1).Selected = True
                        'record position
                        frmH.txtcbxMDBSelIndex.Text = Count1
                        boolFormLoad = boolTTT
                        Exit For
                    End If
                Next

                boolNewOracle = False

            Else
                strM = "The selected study is already configured in StudyDoc."
                strM = strM & ChrW(10) & ChrW(10) & "Please select a different study or click the Cancel button."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                GoTo end1

            End If


            'If rows.Length = 0 Then

            '    'must add this to dgvwstudy
            '    Dim tbl2 As DataTable = tblwSTUDY
            '    Dim dr1 As DataRow = tblwSTUDY.NewRow

            '    dr1.BeginEdit()
            '    dr1("PROJECTIDTEXT") = dgvP("PROJECTIDTEXT", intRowP).Value ' dgvP(intRowP).item("PROJECTIDTEXT")
            '    dr1("STUDYNAME") = dgvS("STUDYNAME", intRowS).Value ' row(0).Item("STUDYNAME")
            '    dr1("STUDYNUMBER") = dgvS("STUDYNUMBER", intRowS).Value 'row(0).Item("STUDYNUMBER")
            '    dr1("SPECIES") = dgvS("SPECIES", intRowS).Value 'row(0).Item("SPECIES")
            '    dr1("STUDYTITLE") = dgvS("STUDYTITLE", intRowS).Value 'row(0).Item("STUDYTITLE")
            '    dr1("PROJECTID") = dgvP("PROJECTID", intRowP).Value ' row(0).Item("PROJECTID")
            '    dr1("STUDYID") = dgvS("STUDYID", intRowS).Value 'row(0).Item("STUDYID")
            '    dr1("SPECIESID") = dgvS("SPECIESID", intRowS).Value 'row(0).Item("SPECIESID")
            '    dr1.EndEdit()

            '    tblwSTUDY.Rows.Add(dr1)


            'Else
            '    strM = "The selected study is already configured in StudyDoc."
            '    strM = strM & ChrW(10) & ChrW(10) & "Please select a different study or click the Cancel button."
            '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            '    GoTo end1
            'End If


            'hmmm. shouldn't need to do this anymore
            GoTo skip1

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

            Dim dt1 As DateTime
            dt1 = Now

            Call frmH.OpenArchiveMDB(False)

skip1:

        Else


        End If



        '*****


        '*****

        boolCancel = False
        Me.Visible = False

end1:

    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Sub LoadGrids()

        Dim dgvP As DataGridView
        Dim var1, var2
        Dim d
        Dim strF As String
        Dim tbl1 As DataTable
        Dim tbl2 As DataTable
        Dim int1 As Int64
        Dim tbl As System.Data.DataTable
        Dim Count1 As Short
        Dim dvP As DataView = New DataView(tblwPROJECTS)


        int1 = tblwPROJECTS.Rows.Count 'debug

        dgvP = Me.dgvProjects

        If boolGetOracle Then
            dgvP.DataSource = dvP 'tblwPROJECTS
        Else
            If boolArchive Then
                Me.dgvProjects.Visible = False
                Me.lblProjects.Visible = False
            Else

                'debug
                For Count1 = 0 To tblwPROJECTS.Columns.Count - 1
                    var1 = tblwPROJECTS.Columns(Count1).ColumnName
                    var1 = var1
                Next

                'LINQ
                Dim qryIDR
                qryIDR = From IDS In tblStudies.AsEnumerable() Join IDR In tblwPROJECTS.AsEnumerable() On IDS("INT_WATSONPROJECTID") Equals IDR("PROJECTID") Select IDS("INT_WATSONPROJECTID")

                tbl2 = tblwSTUDY.Copy

                Try
                    For Each d In qryIDR

                        var1 = d 'this is INT_WATSONID
                        strF = "STUDYID = " & var1
                        Dim row() As DataRow
                        row = tblwSTUDY.Select(strF)

                        If row.Length > 0 Then
                            Dim dr1 As DataRow = tbl2.NewRow

                            dr1.BeginEdit()
                            dr1("PROJECTIDTEXT") = row(0).Item("PROJECTIDTEXT")
                            dr1("STUDYNAME") = row(0).Item("STUDYNAME")
                            dr1("STUDYNUMBER") = row(0).Item("STUDYNUMBER")
                            dr1("SPECIES") = row(0).Item("SPECIES")
                            dr1("STUDYTITLE") = row(0).Item("STUDYTITLE")
                            dr1("PROJECTID") = row(0).Item("PROJECTID")
                            dr1("STUDYID") = row(0).Item("STUDYID")
                            dr1("SPECIESID") = row(0).Item("SPECIESID")
                            dr1.EndEdit()

                            tbl2.Rows.Add(dr1)
                        End If

                    Next

                    'now make unique table
                    Dim dv As DataView
                    dv = New DataView(tbl2)

                    tbl = dv.ToTable("aa", True, "PROJECTID", "PROJECTIDTEXT")

                    dvP = New DataView(tbl)

                Catch ex As Exception
                    var1 = ex.Message
                    MsgBox(ex.Message, vbInformation, "Problem...")
                End Try

                dgvP.DataSource = dvP ' tbl

            End If

        End If

        If dgvP.RowCount = 0 Then
        Else
            dgvP.CurrentCell = dgvP.Item("PROJECTIDTEXT", Count1)
            dgvP.Rows(0).Selected = True
            Call ClickProjects()
        End If

    End Sub

    Sub ClickProjects()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim dgvS As DataGridView

        dgv = Me.dgvProjects
        dgvS = Me.dgvStudies

        Dim intRow As Int32

        If IsNothing(dgv.CurrentRow) Then
            GoTo end1
        End If

        intRow = dgv.CurrentRow.Index

        Dim ID As Int64
        Dim strF As String
        Dim strS As String
        Dim rows() As DataRow
        Dim tbl2 As DataTable
        Dim var1
        Dim dvP As DataView
        Dim dvS As DataView

        If intRow >= 0 Then
            ID = dgv("PROJECTID", intRow).Value
        Else
            ID = 0
        End If

        Dim strSF As String
        strSF = Me.txtFilter.Text

        If Len(strSF) = 0 Then
            strF = "PROJECTID = " & ID
        Else
            strF = "PROJECTID = " & ID & " AND STUDYNAME LIKE '*" & strSF & "*'"
        End If

        Me.txtProjectID.Text = ID.ToString

        Dim tbl1 As DataTable
        rows = tblwSTUDY.Select(strF)
        'boolHold = True
        'dgvS.DataSource = rows
        'boolHold = False

        If rows.Length = 0 Then
            Try
                dgvS.DataSource = Nothing
            Catch ex As Exception
                var1 = ex.Message
            End Try

        Else
            tbl2 = rows.CopyToDataTable
            dvS = New DataView(tbl2)
            boolHold = True
            dgvS.DataSource = dvS ' tbl2
            boolHold = False
        End If


        If dgvS.RowCount = 0 Then
            Me.txtStudyID.Text = 0
        Else
            boolHold = True
            dgvS.Rows(0).Selected = True
            boolHold = False
            Try
                If boolGetOracle Then
                    Me.txtStudyID.Text = dgvS("STUDYID", 0).Value
                Else
                    Try
                        Me.txtStudyID.Text = dgvS("ID_TBLSTUDIES", 0).Value
                    Catch ex As Exception
                        Me.txtStudyID.Text = dgvS("STUDYID", 0).Value
                    End Try

                End If
            Catch ex As Exception
                MsgBox(ex.Message, vbInformation, "Click Projects...")
            End Try
        End If

        Call ConfigGrids(False)

end1:

    End Sub

    Sub ConfigGrids(boolDoProj As Boolean)

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim boolD As Boolean
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim var1
        Dim boolV As Boolean

        Try
            dgv = Me.dgvProjects

            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            If boolDoProj Then

                With dgv.ColumnHeadersDefaultCellStyle
                    .Font = New Font(dgv.Font, FontStyle.Bold)
                End With

                For Count1 = 0 To dgv.ColumnCount - 1
                    str1 = UCase(dgv.Columns(Count1).Name)
                    Select Case str1
                        Case "PROJECTIDTEXT", "PROJECTNAME", "PROJECTTITLE"
                            Select Case str1
                                'Case "PROJECTID"
                                '    str2 = "Project ID"
                                Case "PROJECTIDTEXT"
                                    str2 = "Project ID Text"
                                Case "PROJECTNAME"
                                    str2 = "Project Name"
                                Case "PROJECTTITLE"
                                    str2 = "Project Title"
                            End Select
                            dgv.Columns(Count1).HeaderText = str2
                            dgv.Columns(Count1).Visible = True
                        Case Else
                            dgv.Columns(Count1).Visible = False
                    End Select

                Next

                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                dgv.AutoResizeColumns()


            End If


            dgv = Me.dgvStudies

            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            With dgv.ColumnHeadersDefaultCellStyle
                .Font = New Font(dgv.Font, FontStyle.Bold)
            End With

            var1 = ""

            For Count1 = 0 To dgv.ColumnCount - 1
                str1 = UCase(dgv.Columns(Count1).Name)
                var1 = var1 & "; " & str1
                Select Case str1
                    Case "SPECIES", "STUDYNAME", "STUDYTITLE"
                        Select Case str1
                            'Case "PROJECTID"
                            '    str2 = "Project ID"
                            '    int1 = 0
                            Case "SPECIES"
                                str2 = "Species"
                                int1 = 1
                            Case "STUDYNAME"
                                str2 = "Study Name"
                                int1 = 2
                            Case "STUDYTITLE"
                                str2 = "Study Title"
                                int1 = 3
                        End Select
                        dgv.Columns(Count1).HeaderText = str2
                        dgv.Columns(Count1).DisplayIndex = int1
                        dgv.Columns(Count1).Visible = True
                    Case Else
                        dgv.Columns(Count1).Visible = False
                End Select

            Next

            'display index has to be hit again
            dgv.Columns("SPECIES").DisplayIndex = 1
            dgv.Columns("STUDYNAME").DisplayIndex = 2
            dgv.Columns("STUDYTITLE").DisplayIndex = 3

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AutoResizeColumns()
        Catch ex As Exception
            var1 = ex.Message
            'MsgBox(var1)
        End Try

        

        var1 = var1


    End Sub

    Private Sub dgvStudies_SelectionChanged(sender As Object, e As System.EventArgs) Handles dgvStudies.SelectionChanged

        If boolHold Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim ID As Int64
        Dim intRow As Int32
        Dim var1

        dgv = Me.dgvStudies

        If dgv.RowCount <= 0 Then
            ID = 0
        Else
            If dgv.CurrentRow Is Nothing Then
                dgv.Rows(0).Selected = True
            End If
            intRow = dgv.CurrentRow.Index
            'intRow = dgv.SelectedRows(0).Index
            Try
                var1 = dgv("STUDYID", intRow).Value
            Catch ex As Exception
                MsgBox(ex.Message)
                'var1 = dgv("ID_TBLSTUDIES", intRow).Value
                var1 = 0
            End Try
            ID = var1

        End If

        Me.txtStudyID.Text = ID.ToString

    End Sub

    Private Sub dgvProjects_SelectionChanged(sender As Object, e As System.EventArgs) Handles dgvProjects.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Call ClickProjects()

    End Sub


    Private Sub dgvProjects_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProjects.CellContentClick

    End Sub


    Private Sub txtFilter_TextChanged(sender As Object, e As EventArgs) Handles txtFilter.TextChanged

        Dim var1, var2
        Dim strF As String
        Dim strFS As String
        Dim strFP As String
        Dim Count1 As Int64
        Dim int1 As Int64
        Dim strM As String

        Dim vE, vD
        vE = Color.Gainsboro ' "Gainsboro"
        vD = Color.DarkGray

        Dim tblPID As New DataTable
        Dim col1 As New DataColumn
        col1.ColumnName = "ProjectID"
        col1.DataType = System.Type.GetType("System.Int64")
        tblPID.Columns.Add(col1)

        Dim dgvP As DataGridView = Me.dgvProjects
        Dim dv As DataView = dgvP.DataSource

        var1 = Me.txtFilter.Text

        If Len(var1) = 0 Then
            dv.RowFilter = ""
            Me.cmdOK.Enabled = True
            Me.cmdOK.BackColor = vE
            GoTo end1
        End If

        strFS = "STUDYNAME LIKE '*" & var1 & "*'"
        Dim rows() As DataRow

        Try
            rows = tblwSTUDY.Select(strFS, "STUDYNAME ASC")
        Catch ex As Exception

            strM = "The filter probably contains an invalid character."
            Me.cmdOK.BackColor = vD
            Me.cmdOK.Enabled = False
            MsgBox(strM, vbInformation, "Invalid action..")
            GoTo end1

        End Try


        If rows.Length = 0 Then
            dv.RowFilter = "PROJECTID = -100"
            Dim dvS As DataView = Me.dgvStudies.DataSource
            'dvS.RowFilter = "STUDYNAME = '" & var1 & "'"

            dvS.RowFilter = "PROJECTID = -100"
            'If IsNothing(dvS) Then
            'Else
            '    dvS.RowFilter = "PROJECTID = -100"
            'End If

            Me.cmdOK.Enabled = False
            Me.cmdOK.BackColor = vD

            GoTo end1
        Else

            Me.cmdOK.Enabled = True

            Dim boolE As Boolean = False
            int1 = 0
            For Count1 = 0 To rows.Length - 1

                If int1 > 600 Then
                    boolE = True
                    Exit For
                End If
                var2 = NZ(rows(Count1).Item("PROJECTID"), "")
                If Len(var2) = 0 Then
                Else
                    'var2 = CLng(var2)
                    If Count1 = 0 Then
                        int1 = int1 + 1
                        strFP = "PROJECTID = " & var2
                        Dim nr As DataRow = tblPID.NewRow
                        nr.BeginEdit()
                        nr.Item("ProjectID") = var2
                        nr.EndEdit()
                        tblPID.Rows.Add(nr)
                    Else

                        strF = "ProjectID = " & var2
                        Dim rowsPID() As DataRow = tblPID.Select(strF)
                        If rowsPID.Length = 0 Then

                            int1 = int1 + 1

                            strFP = strFP & " OR PROJECTID = " & var2
                            Dim nr As DataRow = tblPID.NewRow
                            nr.BeginEdit()
                            nr.Item("ProjectID") = var2
                            nr.EndEdit()
                            tblPID.Rows.Add(nr)
                        End If
                    End If
                End If

            Next

            var1 = tblPID.Rows.Count 'evaluate number of records

            Try
                dv.RowFilter = strFP
                If dv.Count = 0 Then
                Else
                    dgvP.CurrentCell = dgvP.Item("PROJECTIDTEXT", 0)
                    dgvP.Rows(0).Selected = True
                    Call ClickProjects()
                End If
            Catch ex As Exception
                boolE = True
                Me.cmdOK.Enabled = False
                Me.cmdOK.BackColor = vD
            End Try


            If boolE Then
                strM = "The filter is too large and the filter is not complete. Please type some more characters."
                MsgBox(strM, vbInformation, "Invalid action..")
            Else
                Me.cmdOK.BackColor = vE
                Me.cmdOK.Enabled = True
            End If

        End If

end1:

    End Sub

    Private Sub cmdClear_Click(sender As Object, e As EventArgs) Handles cmdClear.Click

        Me.txtFilter.Clear()

    End Sub

    Private Sub dgvStudies_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvStudies.CellContentClick

    End Sub
End Class