
Imports System.Data
Imports System.Linq.Expressions
Imports System.Linq
Imports System.Data.DataTableExtensions
Imports System.Data.DataRowExtensions

Public Class frmClearStudy

    Public boolCancel As Boolean
    Public boolHold As Boolean = False
    Public boolGetOracle As Boolean
    Public boolArchive As Boolean
    Public boolCompare As Boolean = False
    Public boolFormLoad As Boolean = False
    Public boolConfig As Boolean = False


    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdApply.Click


        Dim intR As Short
        Dim strM As String
        Dim str1 As String
        Dim str2 As String
        Dim strReason As String = Me.txtReason.Text

        'validate
        If Len(strReason) = 0 Then
            strM = "Reason for deletion cannot be blank"
            MsgBox(strM, vbInformation, "Invalid action...")
            Exit Sub
        End If



        Try

            Dim dgv As DataGridView = Me.dgvStudies
            Dim rows As DataGridViewSelectedRowCollection = dgv.SelectedRows

            If rows.Count = 0 Then
                strM = "A study must be chosen."
                str2 = "Invalid action..."
                MsgBox(strM, vbInformation, str2)
                GoTo end1
            End If

            Dim id_S As Int64
            Dim strSN As String 'study name

            Dim dr As DataGridViewRow = rows(0)
            str1 = dr.Cells("CHARWATSONSTUDYNAME").Value
            id_S = dr.Cells("ID_TBLSTUDIES").Value
            strSN = dr.Cells("CHARWATSONSTUDYNAME").Value

            'check to see if id_S is current study
            If id_S = id_tblStudies Then
                strM = "Cannot clear the current study from the StudyDoc database." & ChrW(10) & ChrW(10)
                strM = strM & "Please choose a different study, then attempt to clear this study."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If

            strM = "Are you sure you wish to clear the study '" & str1 & "' from the StudyDoc database?"
            strM = strM & ChrW(10) & ChrW(10) & "After this action is completed, the first study in the list will be selected and loaded." & ChrW(10) & ChrW(10)
            strM = strM & "If enabled, an Audit Trail entry will be recorded describing that the study has been cleared from the StudyDoc database."
            str2 = "Continue?"
            intR = MsgBox(strM, vbOKCancel, str2)

            If intR = 1 Then
            Else
                GoTo end1
            End If

            Call ClearStudy(id_S, strSN, strReason)

            Me.Visible = False

        Catch ex As Exception

            strM = "There was a problem clearing the study '" & str1 & "' from the StudyDoc database?" & ChrW(10) & ChrW(10)
            strM = strM & ex.Message
            MsgBox(strM, vbInformation, "Problem...")

        End Try


end1:


    End Sub


    Sub ClearStudy(ByVal id_S As Int64, ByVal strStudyName As String, ByVal strReason As String)

        '20171110 LEE: Need a function to clear study from StudyDoc database
        'at this point in time, restricted to only Watson = Oracle

        Dim Count1 As Short
        Dim Count2 As Int64

        Dim strF As String
        Dim strT As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim var1, var2, var3
        Dim strM As String
        Dim intR As Int16

        'do this item in order to trigger an audit trail item
        If gboolAuditTrail Then

            'first delete all entries in tblAuditTrailTemp
            tblAuditTrailTemp.Clear()
            tblAuditTrailTemp.AcceptChanges()

            Dim rows() As DataRow = tblStudies.Select("ID_TBLSTUDIES = " & id_S)
            gIDDeleteStudy = id_S 'need this in FillAuditTrailTemp

            If rows.Length = 0 Then
            Else
                rows(0).Delete()
            End If

            'enter and audit trail item
            Call FillAuditTrailTemp(tblStudies)

        End If


        'first delete from some tables for current information
        'Nope, don't do this
        'has to go through data adapter
        'Dim dvTemp1 As DataView
        'strF = "ID_TBLSTUDIES <> " & id_S

        'dvTemp1 = tblStudies.DefaultView
        'dvTemp1.RowFilter = strF
        'tblStudies = dvTemp1.ToTable

        'dvTemp1 = tblData.DefaultView
        'dvTemp1.RowFilter = strF
        'tblData = dvTemp1.ToTable

        'dvTemp1 = tblReports.DefaultView
        'dvTemp1.RowFilter = strF
        'tblReports = dvTemp1.ToTable


        Dim boolProb As Boolean = False
        Dim constr As String = constrIni

        Dim myConnection As OleDb.OleDbConnection
        Dim cmd As New OleDb.OleDbCommand

        Try

            myConnection = New OleDb.OleDbConnection(constrIni)

            myConnection.Open()
            cmd.Connection = myConnection
            cmd.CommandType = CommandType.Text

        Catch ex As Exception
            var1 = ex.Message
        End Try

        Dim boolSkip As Boolean = False

        For Count1 = 1 To 38

            boolSkip = False

            Select Case Count1
                Case 1
                    strT = "TBLANALREFSTANDARDS"
                Case 2
                    strT = "TBLANALYTICALRUNSUMMARY"
                Case 3
                    strT = "TBLAPPFIGS"
                Case 4
                    strT = "TBLAPPFIGWORDDOCS"
                Case 5
                    strT = "TBLASSIGNEDSAMPLES"
                Case 6
                    strT = "TBLAUDITTRAIL"
                    boolSkip = True
                Case 7
                    strT = "TBLAUTOASSIGNSAMPLES"
                Case 8
                    strT = "TBLCONTRIBUTINGPERSONNEL"
                Case 9
                    strT = "TBLCUSTOMFIELDCODES"
                Case 10
                    strT = "TBLDATA"
                Case 11
                    strT = "TBLFINALREPORT"
                Case 12
                    strT = "TBLFINALREPORTWORDDOCS"
                Case 13
                    strT = "TBLGUWUASSAY"
                Case 14
                    strT = "TBLGUWUASSAYPERS"
                Case 15
                    strT = "TBLGUWUASSIGNEDCMPD"
                Case 16
                    strT = "TBLGUWUASSIGNEDCMPDLOT"
                Case 17
                    strT = "TBLGUWUPKSUBJECTS"
                Case 18
                    strT = "TBLGUWURTTIMEPOINTS"
                Case 19
                    strT = "TBLGUWUSTUDIES"
                Case 20
                    strT = "TBLINCLUDEDROWS"
                Case 21
                    strT = "TBLMETHODVALIDATIONDATA"
                Case 22
                    strT = "TBLOUTSTANDINGITEMS"
                Case 23
                    strT = "TBLQATABLES"
                Case 24
                    strT = "TBLREPORTHEADERS"
                Case 25
                    strT = "TBLREPORTHISTORY"
                Case 26
                    strT = "TBLREPORTS"
                Case 27
                    strT = "TBLREPORTSTATEMENTS"
                Case 28
                    strT = "TBLREPORTTABLE"
                Case 29
                    strT = "TBLREPORTTABLEANALYTES"
                Case 30
                    strT = "TBLREPORTTABLEHEADERCONFIG"
                Case 31
                    strT = "TBLSAMPLERECEIPT"
                Case 32
                    strT = "TBLSAVEEVENT"
                Case 33
                    strT = "TBLSTUDIES"
                Case 34
                    strT = "TBLSTUDYDOCANALYTES"
                Case 35
                    strT = "TBLSUMMARYDATA"
                Case 36
                    strT = "TBLTABLELEGENDS"
                Case 37
                    strT = "TBLTABLEPROPERTIES"
                Case 38
                    strT = "TBLTEMPLATES"
            End Select

            If boolSkip Then
            Else
                str1 = "DELETE FROM " & strT & " "
                str2 = "WHERE ID_TBLSTUDIES = " & id_S & ";"
                strSQL = str1 & str2

                cmd.CommandText = strSQL

                Try
                    cmd.ExecuteNonQuery()
                    boolProb = False
                    strM = """" & strSQL & """ successfully executed on " & strT & "." '& vbCrLf

                Catch ex As Exception
                    boolProb = True
                    strM = "There was a problem executing """ & strSQL & """ on " & strT & ". Please contact your LABIntegrity StudyDoc technical representative." & ChrW(10) & ChrW(10) & ex.Message
                    MsgBox(Count1 & ":  " & strM)
                End Try
            End If

        Next Count1

        'now re-establish some StudyDoc tables
        'tblStudies
        'tblData
        'tblReports

        'do this item in order to trigger an audit trail item
        If gboolAuditTrail Then

            'first delete all entries in tblAuditTrailTemp
            tblAuditTrailTemp.Clear()
            tblAuditTrailTemp.AcceptChanges()

            Dim rows() As DataRow = tblStudies.Select("ID_TBLSTUDIES = " & id_S)
            gIDDeleteStudy = id_S 'need this in FillAuditTrailTemp

            If rows.Length = 0 Then
            Else
                rows(0).Delete()
            End If

            'enter an audit trail item
            Call FillAuditTrailTemp(tblStudies)

        End If



        Dim boolRefreshA As Boolean

        If boolGuWuOracle Then
            boolRefreshA = DAsRefresh(frmH)
        ElseIf boolGuWuAccess Then
            boolRefreshA = DAsRefreshAcc(frmH)
        ElseIf boolGuWuSQLServer Then
            boolRefreshA = DAsRefreshSQLServer(frmH)
        End If

        frmH.lblProgress.Visible = False
        frmH.panProgress.Visible = False
        frmH.Refresh()

        'now re-establis dgvwStudy on Home
        '
        Dim boolW As Boolean

        If boolAccess Then
            boolW = False 'Watson is .mdb
        Else
            boolW = True 'Watson is Oracle
        End If

        'open connection
        'constrCur comes from frmHome.Load
        If boolW Then

            Dim wcn As New ADODB.Connection

            wcn.Open(constrCur)

            Try
                Call Configure_dgvwStudy(boolW, wcn, boolANSI)
                If boolW Then
                    Try
                        Call ConfigStudyTable(True, True)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If

            Catch ex As Exception

            End Try

            If boolW Then

            End If

            'close connection
            Try
                wcn.Close()
            Catch ex As Exception
                var1 = var1 'debug
            End Try

            wcn = Nothing

        End If

        If gboolAuditTrail Then

            Dim strRFCTemp As String = strRFC
            strRFC = strReason

            'record audit trail
            Call RecordAuditTrail(False, Now)

            strRFC = strRFCTemp

        End If


end1:


    End Sub

    Private Sub frmClearStudy_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Call ControlDefaults(Me)

        Call DoubleBufferControl(Me, "dgv")

        boolFormLoad = True

        'Call ConfigStudies()

        Try
            Call LoadGrids()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Call ConfigGrids(True)

        boolFormLoad = False

        If Me.dgvProjects.RowCount = 0 Then
        Else
            Me.dgvProjects.CurrentCell = Me.dgvProjects.Item("PROJECTIDTEXT", 0)
            Me.dgvProjects.Rows(0).Selected = True
            Call ClickProjects()
        End If

        Try
            Me.txtFilter.Focus()
        Catch ex As Exception

        End Try

    End Sub

    Sub LoadGrids()

        Dim dgvP As DataGridView = Me.dgvProjects
        Dim dgvS As DataGridView = Me.dgvStudies
        Dim var1, var2
        Dim a, b, c, d
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
            If boolAccess Then
                Me.dgvProjects.Visible = False
                Me.lblProjects.Visible = False

                Me.panStudies.Left = Me.lblProjects.Left

                a = dgvS.Left + dgvS.Width
                dgvS.Left = dgvP.Left
                dgvS.Width = a - dgvS.Left

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

        If boolAccess Then
            Call ConfigStudies()
        Else
            If dgvP.RowCount = 0 Then
            Else
                dgvP.CurrentCell = dgvP.Item("PROJECTIDTEXT", 0)
                dgvP.Rows(0).Selected = True
                Call ClickProjects()
            End If
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

        Call ConfigStudies()

        GoTo end1

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

end1:

        Call ConfigGrids(False)



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

                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                dgv.AutoResizeColumns()
                dgv.AutoResizeRows()


            End If

            GoTo end1


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


end1:

        var1 = var1


    End Sub

    Sub ConfigStudies()

        Dim dgv As DataGridView = Me.dgvStudies

        Dim Count1 As Integer
        Dim strF As String
        Dim strS As String
        Dim int1 As Int64

        Dim idP As Int64

        If boolAccess Then
            strF = "INT_WATSONSTUDYID > 0"
        Else
            idP = Me.txtProjectID.Text
            strF = "INT_WATSONSTUDYID > 0 AND INT_WATSONPROJECTID = " & idP
        End If

        strS = "CHARWATSONSTUDYNAME ASC"

        Dim dv As DataView = New DataView(tblStudies, strF, strS, DataViewRowState.CurrentRows)

        int1 = dv.Count 'debug

        dgv.DataSource = dv

        With dgv.ColumnHeadersDefaultCellStyle
            .Font = New Font(dgv.Font, FontStyle.Bold)
        End With

        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        dgv.Columns("CHARWATSONSTUDYNAME").Visible = True
        dgv.Columns("CHARWATSONSTUDYNAME").HeaderText = "Watson Study Name"
        dgv.Columns("CHARWATSONSTUDYNAME").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        If boolConfig Then
        Else
            'dgv.RowHeadersWidth = dgv.RowHeadersWidth * 0.5
            boolConfig = True
        End If

        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        Me.Visible = False

    End Sub

    Private Sub dgvProjects_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProjects.CellContentClick

    End Sub

    Private Sub dgvProjects_SelectionChanged(sender As Object, e As EventArgs) Handles dgvProjects.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Call ClickProjects()

    End Sub

    Private Sub txtFilter_TextChanged(sender As Object, e As EventArgs) Handles txtFilter.TextChanged


        Dim var1, var2
        Dim strF As String
        Dim strFS As String
        Dim strFP As String
        Dim Count1 As Int64
        Dim Count2 As Int32
        Dim int1 As Int64
        Dim strM As String
        Dim strFind As String
        Dim str1 As String
        Dim str2 As String

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
        strFind = var1

        If Len(var1) = 0 Then
            If boolAccess Then

            Else
                dv.RowFilter = ""
            End If

            Me.cmdApply.Enabled = True
            Me.cmdApply.BackColor = vE
            GoTo end1
        End If

        strFS = "CHARWATSONSTUDYNAME LIKE '*" & var1 & "*'"
        Dim rows() As DataRow

        Try
            rows = tblStudies.Select(strFS, "CHARWATSONSTUDYNAME ASC")
        Catch ex As Exception

            strM = "The filter probably contains an invalid character."
            Me.cmdApply.BackColor = vD
            Me.cmdApply.Enabled = False
            MsgBox(strM, vbInformation, "Invalid action..")
            GoTo end1

        End Try

        Dim dvS As DataView = Me.dgvStudies.DataSource
        Dim sss As String
        sss = dvS.Sort
        var1 = var1

        If rows.Length = 0 Then

            dv.RowFilter = "PROJECTID = -100"

            'dvS.RowFilter = "STUDYNAME = '" & var1 & "'"

            dvS.RowFilter = "INT_WATSONPROJECTID = -100"
            'If IsNothing(dvS) Then
            'Else
            '    dvS.RowFilter = "PROJECTID = -100"
            'End If

            Me.cmdApply.Enabled = False
            Me.cmdApply.BackColor = vD

            GoTo end1

        Else

            Me.cmdApply.Enabled = True

            Dim boolE As Boolean = False
            int1 = 0

            If boolAccess Then
                dvS.RowFilter = strFS
                dgvStudies.AutoResizeRows()
            Else
                For Count1 = 0 To rows.Length - 1

                    If int1 > 600 Then
                        boolE = True
                        Exit For
                    End If
                    var2 = NZ(rows(Count1).Item("INT_WATSONPROJECTID"), "")
                    If Len(var2) = 0 Then
                    Else
                        'var2 = CLng(var2)
                        If Count1 = 0 Then
                            int1 = int1 + 1
                            strFP = "PROJECTID = " & var2
                            Dim nr As DataRow = tblPID.NewRow
                            nr.BeginEdit()
                            nr.Item("PROJECTID") = var2
                            nr.EndEdit()
                            tblPID.Rows.Add(nr)
                        Else

                            strF = "PROJECTID = " & var2
                            Dim rowsPID() As DataRow = tblPID.Select(strF)
                            If rowsPID.Length = 0 Then

                                int1 = int1 + 1

                                strFP = strFP & " OR PROJECTID = " & var2
                                Dim nr As DataRow = tblPID.NewRow
                                nr.BeginEdit()
                                nr.Item("PROJECTID") = var2
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
                    Me.cmdApply.Enabled = False
                    Me.cmdApply.BackColor = vD
                End Try

            End If

            If boolE Then
                strM = "The filter is too large and the filter is not complete. Please type some more characters."
                MsgBox(strM, vbInformation, "Invalid action..")
            Else
                Me.cmdApply.BackColor = vE
                Me.cmdApply.Enabled = True
            End If

        End If

        'check to see if item is in list
        For Count2 = 0 To dgvStudies.Rows.Count - 1
            str1 = dgvStudies("CHARWATSONSTUDYNAME", Count2).Value
            If InStr(1, str1, strFind, CompareMethod.Text) > 0 Then
                Me.dgvStudies.CurrentCell = Me.dgvStudies.Item("CHARWATSONSTUDYNAME", Count2)
                Exit For
            End If
        Next

end1:

    End Sub

    Private Sub cmdClear_Click(sender As Object, e As EventArgs) Handles cmdClear.Click

        Me.txtFilter.Clear()

    End Sub
End Class